package main

import (
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

// GeocodeResponse represents the response from Nominatim API
type GeocodeResponse struct {
	DisplayName string `json:"display_name"`
	Address     struct {
		HouseNumber   string `json:"house_number"`
		Road          string `json:"road"`
		Suburb        string `json:"suburb"`
		City          string `json:"city"`
		County        string `json:"county"`
		State         string `json:"state"`
		StateDistrict string `json:"state_district"`
		Postcode      string `json:"postcode"`
		Country       string `json:"country"`
		// Thailand specific fields
		Subdistrict string `json:"subdistrict"`
		District    string `json:"district"`
		Province    string `json:"province"`
	} `json:"address"`
}

// Coordinates represents latitude and longitude
type Coordinates struct {
	Lat float64
	Lng float64
}

// Repository handles Excel file operations
type Repository struct {
	file      *excelize.File
	sheetName string
	rows      [][]string
}

// NewRepository creates a new repository instance
func NewRepository(excelFile string) (*Repository, error) {
	f, err := excelize.OpenFile(excelFile)
	if err != nil {
		return nil, fmt.Errorf("opening Excel file: %w", err)
	}

	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		f.Close()
		return nil, fmt.Errorf("no sheets found in Excel file")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		f.Close()
		return nil, fmt.Errorf("reading rows: %w", err)
	}

	if len(rows) == 0 {
		f.Close()
		return nil, fmt.Errorf("Excel file is empty")
	}

	return &Repository{
		file:      f,
		sheetName: sheetName,
		rows:      rows,
	}, nil
}

// Close closes the Excel file
func (r *Repository) Close() error {
	return r.file.Close()
}

// GetSheetName returns the sheet name
func (r *Repository) GetSheetName() string {
	return r.sheetName
}

// GetRows returns all rows
func (r *Repository) GetRows() [][]string {
	return r.rows
}

// GetFile returns the Excel file handle
func (r *Repository) GetFile() *excelize.File {
	return r.file
}

// SaveAs saves the file to the specified path
func (r *Repository) SaveAs(outputFile string) error {
	return r.file.SaveAs(outputFile)
}

// SetCellValue sets a cell value
func (r *Repository) SetCellValue(cell string, value interface{}) error {
	return r.file.SetCellValue(r.sheetName, cell, value)
}

// Service handles business logic for coordinate to address conversion
type Service struct {
	repo  *Repository
	cache *coordinateCache
}

// NewService creates a new service instance
func NewService(repo *Repository) *Service {
	return &Service{
		repo:  repo,
		cache: newCoordinateCache(),
	}
}

// Process converts coordinates to addresses and saves the result
func (s *Service) Process(excelFile string) error {
	fmt.Printf("Processing sheet: %s\n", s.repo.GetSheetName())

	rows := s.repo.GetRows()
	totalRows := len(rows) - 1 // Exclude header
	fmt.Printf("Total rows to process: %d\n", totalRows)

	latLngCol, addressCol, districtCol, provinceCol, err := s.findColumns(rows)
	if err != nil {
		return err
	}

	if addressCol == -1 || districtCol == -1 || provinceCol == -1 {
		addressCol, districtCol, provinceCol = s.addAddressColumns(len(rows[0]))
	}

	// For large datasets (>100k rows), process in batches and save periodically
	batchSize := 1000
	if totalRows > 100000 {
		fmt.Printf("Large dataset detected. Processing in batches of %d rows...\n", batchSize)
		processed := s.processRowsInBatches(rows, latLngCol, addressCol, districtCol, provinceCol, batchSize, excelFile)
		fmt.Printf("\n✓ Processed %d rows\n", processed)
	} else {
		processed := s.processRows(rows, latLngCol, addressCol, districtCol, provinceCol)
		fmt.Printf("\n✓ Processed %d rows\n", processed)
	}

	// Save to data/ directory
	dataDir := "data"
	if err := os.MkdirAll(dataDir, 0755); err != nil {
		return fmt.Errorf("creating data directory: %w", err)
	}

	fileName := filepath.Base(excelFile)
	outputFile := filepath.Join(dataDir, strings.TrimSuffix(fileName, ".xlsx")+"_with_addresses.xlsx")
	if err := s.repo.SaveAs(outputFile); err != nil {
		return fmt.Errorf("saving file: %w", err)
	}

	fmt.Printf("✓ Output saved to: %s\n", outputFile)
	return nil
}

// processRowsInBatches processes rows in batches for large datasets
func (s *Service) processRowsInBatches(rows [][]string, latLngCol, addressCol, districtCol, provinceCol, batchSize int, excelFile string) int {
	totalRows := len(rows) - 1
	totalBatches := (totalRows + batchSize - 1) / batchSize
	processed := 0

	for batch := 0; batch < totalBatches; batch++ {
		start := batch*batchSize + 1 // +1 to skip header
		end := start + batchSize
		if end > len(rows) {
			end = len(rows)
		}

		fmt.Printf("\n--- Processing batch %d/%d (rows %d-%d) ---\n", batch+1, totalBatches, start, end-1)

		// Process this batch
		batchRows := rows[start:end]
		// Adjust row indices for batch processing
		batchProcessed := s.processBatch(batchRows, start-1, latLngCol, addressCol, districtCol, provinceCol)
		processed += batchProcessed

		// Save progress after each batch
		dataDir := "data"
		fileName := filepath.Base(excelFile)
		tempFile := filepath.Join(dataDir, strings.TrimSuffix(fileName, ".xlsx")+"_temp.xlsx")
		if err := s.repo.SaveAs(tempFile); err != nil {
			fmt.Printf("Warning: Could not save progress: %v\n", err)
		} else {
			fmt.Printf("Progress saved: %d/%d rows processed (%.1f%%)\n", processed, totalRows, float64(processed)/float64(totalRows)*100)
		}

		// Small delay between batches to be respectful
		time.Sleep(500 * time.Millisecond)
	}

	return processed
}

// processBatch processes a batch of rows
func (s *Service) processBatch(batchRows [][]string, startIndex, latLngCol, addressCol, districtCol, provinceCol int) int {
	numWorkers := 10
	requestDelay := 1500 * time.Millisecond

	jobs := make(chan int, len(batchRows))
	results := make(chan rowResult, len(batchRows))
	var wg sync.WaitGroup

	// Start workers
	for w := 0; w < numWorkers; w++ {
		wg.Add(1)
		go func(workerID int) {
			defer wg.Done()
			for batchIdx := range jobs {
				rowIndex := startIndex + batchIdx
				row := batchRows[batchIdx]

				// Ensure row has enough columns
				maxCol := latLngCol
				if addressCol > maxCol {
					maxCol = addressCol
				}
				if districtCol > maxCol {
					maxCol = districtCol
				}
				if provinceCol > maxCol {
					maxCol = provinceCol
				}
				for len(row) <= maxCol {
					row = append(row, "")
				}

				coordStr := strings.TrimSpace(row[latLngCol])
				if coordStr == "" {
					results <- rowResult{rowIndex: rowIndex, skipped: true, message: "empty coordinates"}
					continue
				}

				coords, err := s.parseCoordinates(coordStr)
				if err != nil {
					results <- rowResult{rowIndex: rowIndex, skipped: true, message: err.Error()}
					continue
				}

				// Check cache first (for duplicate coordinates)
				address, district, province, cached := s.cache.get(coords.Lat, coords.Lng)
				if !cached {
					// Rate limiting per worker
					time.Sleep(requestDelay)

					address, district, province, err = s.reverseGeocode(coords.Lat, coords.Lng)
					if err != nil {
						results <- rowResult{
							rowIndex: rowIndex,
							skipped:  true,
							message:  fmt.Sprintf("geocode error: %v", err),
						}
						continue
					}

					// Cache the result
					s.cache.set(coords.Lat, coords.Lng, address, district, province)
				}

				results <- rowResult{
					rowIndex: rowIndex,
					address:  address,
					district: district,
					province: province,
					coords:   coords,
				}
			}
		}(w)
	}

	// Send jobs
	go func() {
		for i := 0; i < len(batchRows); i++ {
			jobs <- i
		}
		close(jobs)
	}()

	// Close results channel when all workers are done
	go func() {
		wg.Wait()
		close(results)
	}()

	// Process results
	batchProcessed := 0
	for result := range results {
		rowNum := result.rowIndex + 1

		if result.skipped {
			if rowNum%100 == 0 || strings.Contains(result.message, "rate limit") {
				fmt.Printf("Row %d: %s\n", rowNum, result.message)
			}
			continue
		}

		// Write full address
		colName, _ := excelize.ColumnNumberToName(addressCol + 1)
		cell := fmt.Sprintf("%s%d", colName, rowNum)
		s.repo.SetCellValue(cell, result.address)

		// Write district
		colName, _ = excelize.ColumnNumberToName(districtCol + 1)
		cell = fmt.Sprintf("%s%d", colName, rowNum)
		s.repo.SetCellValue(cell, result.district)

		// Write province
		colName, _ = excelize.ColumnNumberToName(provinceCol + 1)
		cell = fmt.Sprintf("%s%d", colName, rowNum)
		s.repo.SetCellValue(cell, result.province)

		batchProcessed++
		if batchProcessed%100 == 0 {
			fmt.Printf("  Processed %d rows in this batch...\n", batchProcessed)
		}
	}

	return batchProcessed
}

// findColumns finds the latitude/longitude, address, district, and province columns
func (s *Service) findColumns(rows [][]string) (latLngCol, addressCol, districtCol, provinceCol int, err error) {
	headerRow := rows[0]
	latLngCol = -1
	addressCol = -1
	districtCol = -1
	provinceCol = -1

	// Check header row
	for i, cell := range headerRow {
		cellLower := strings.ToLower(strings.TrimSpace(cell))
		if latLngCol == -1 && (strings.Contains(cellLower, "latlg") ||
			strings.Contains(cellLower, "lat") ||
			strings.Contains(cellLower, "coordinate") ||
			strings.Contains(cellLower, "coord")) {
			latLngCol = i
		}
		if strings.Contains(cellLower, "address") {
			addressCol = i
		}
		if strings.Contains(cellLower, "district") {
			districtCol = i
		}
		if strings.Contains(cellLower, "province") {
			provinceCol = i
		}
	}

	// If not found in header, check first data row for comma-separated format
	if latLngCol == -1 && len(rows) > 1 {
		latLngCol = s.detectCoordinateColumn(rows[1])
	}

	if latLngCol == -1 {
		return -1, -1, -1, -1, fmt.Errorf("could not find latitude/longitude column. Please ensure your Excel file has a column with coordinates in format 'lat,lng' (e.g., '13.536964,105.927722') or a header containing 'latlg', 'lat', or 'coordinate'")
	}

	fmt.Printf("Found coordinates column: %s (column %d)\n", headerRow[latLngCol], latLngCol+1)
	return latLngCol, addressCol, districtCol, provinceCol, nil
}

// detectCoordinateColumn detects coordinate column by checking for comma-separated numbers
func (s *Service) detectCoordinateColumn(row []string) int {
	for i, cell := range row {
		if strings.Contains(cell, ",") {
			parts := strings.Split(cell, ",")
			if len(parts) == 2 {
				if _, err1 := strconv.ParseFloat(strings.TrimSpace(parts[0]), 64); err1 == nil {
					if _, err2 := strconv.ParseFloat(strings.TrimSpace(parts[1]), 64); err2 == nil {
						return i
					}
				}
			}
		}
	}
	return -1
}

// addAddressColumns adds Address, District, and Province columns to the Excel file
func (s *Service) addAddressColumns(currentColCount int) (addressCol, districtCol, provinceCol int) {
	addressCol = currentColCount
	colName, _ := excelize.ColumnNumberToName(addressCol + 1)
	s.repo.SetCellValue(fmt.Sprintf("%s1", colName), "Address")
	fmt.Printf("Added Address column at column %d\n", addressCol+1)

	districtCol = currentColCount + 1
	colName, _ = excelize.ColumnNumberToName(districtCol + 1)
	s.repo.SetCellValue(fmt.Sprintf("%s1", colName), "District")
	fmt.Printf("Added District column at column %d\n", districtCol+1)

	provinceCol = currentColCount + 2
	colName, _ = excelize.ColumnNumberToName(provinceCol + 1)
	s.repo.SetCellValue(fmt.Sprintf("%s1", colName), "Province")
	fmt.Printf("Added Province column at column %d\n", provinceCol+1)

	return addressCol, districtCol, provinceCol
}

// parseCoordinates parses comma-separated coordinates string
func (s *Service) parseCoordinates(coordStr string) (Coordinates, error) {
	parts := strings.Split(coordStr, ",")
	if len(parts) != 2 {
		return Coordinates{}, fmt.Errorf("invalid format, expected 'lat,lng'")
	}

	lat, err := strconv.ParseFloat(strings.TrimSpace(parts[0]), 64)
	if err != nil {
		return Coordinates{}, fmt.Errorf("invalid latitude: %w", err)
	}

	lng, err := strconv.ParseFloat(strings.TrimSpace(parts[1]), 64)
	if err != nil {
		return Coordinates{}, fmt.Errorf("invalid longitude: %w", err)
	}

	return Coordinates{Lat: lat, Lng: lng}, nil
}

// rowResult holds the result of processing a row
type rowResult struct {
	rowIndex int
	skipped  bool
	message  string
	address  string
	district string
	province string
	coords   Coordinates
}

// coordinateCache caches geocoding results to avoid duplicate API calls
type coordinateCache struct {
	mu    sync.RWMutex
	cache map[string]cacheEntry
}

type cacheEntry struct {
	address  string
	district string
	province string
}

func newCoordinateCache() *coordinateCache {
	return &coordinateCache{
		cache: make(map[string]cacheEntry),
	}
}

func (c *coordinateCache) get(lat, lng float64) (address, district, province string, found bool) {
	key := fmt.Sprintf("%.6f,%.6f", lat, lng)
	c.mu.RLock()
	defer c.mu.RUnlock()
	entry, exists := c.cache[key]
	if exists {
		return entry.address, entry.district, entry.province, true
	}
	return "", "", "", false
}

func (c *coordinateCache) set(lat, lng float64, address, district, province string) {
	key := fmt.Sprintf("%.6f,%.6f", lat, lng)
	c.mu.Lock()
	defer c.mu.Unlock()
	c.cache[key] = cacheEntry{
		address:  address,
		district: district,
		province: province,
	}
}

// processRows processes all data rows and converts coordinates to addresses concurrently
func (s *Service) processRows(rows [][]string, latLngCol, addressCol, districtCol, provinceCol int) int {
	// Number of concurrent workers (10 workers for faster processing)
	numWorkers := 10
	// Rate limiting: delay between requests per worker (1.5 seconds per worker)
	requestDelay := 1500 * time.Millisecond

	// Channel for jobs
	jobs := make(chan int, len(rows))
	results := make(chan rowResult, len(rows))
	var wg sync.WaitGroup

	// Start workers
	for w := 0; w < numWorkers; w++ {
		wg.Add(1)
		go func(workerID int) {
			defer wg.Done()
			for rowIndex := range jobs {
				row := rows[rowIndex]

				// Ensure row has enough columns
				maxCol := latLngCol
				if addressCol > maxCol {
					maxCol = addressCol
				}
				if districtCol > maxCol {
					maxCol = districtCol
				}
				if provinceCol > maxCol {
					maxCol = provinceCol
				}
				for len(row) <= maxCol {
					row = append(row, "")
				}

				coordStr := strings.TrimSpace(row[latLngCol])
				if coordStr == "" {
					results <- rowResult{rowIndex: rowIndex, skipped: true, message: "empty coordinates"}
					continue
				}

				coords, err := s.parseCoordinates(coordStr)
				if err != nil {
					results <- rowResult{rowIndex: rowIndex, skipped: true, message: err.Error()}
					continue
				}

				// Check cache first (for duplicate coordinates)
				address, district, province, cached := s.cache.get(coords.Lat, coords.Lng)
				if !cached {
					// Rate limiting per worker
					time.Sleep(requestDelay)

					address, district, province, err = s.reverseGeocode(coords.Lat, coords.Lng)
					if err != nil {
						results <- rowResult{
							rowIndex: rowIndex,
							skipped:  true,
							message:  fmt.Sprintf("geocode error: %v", err),
						}
						continue
					}

					// Cache the result
					s.cache.set(coords.Lat, coords.Lng, address, district, province)
				}

				results <- rowResult{
					rowIndex: rowIndex,
					address:  address,
					district: district,
					province: province,
					coords:   coords,
				}
			}
		}(w)
	}

	// Send jobs
	go func() {
		for i := 1; i < len(rows); i++ {
			jobs <- i
		}
		close(jobs)
	}()

	// Close results channel when all workers are done
	go func() {
		wg.Wait()
		close(results)
	}()

	// Process results
	processed := 0
	completed := 0
	total := len(rows) - 1

	for result := range results {
		completed++
		rowNum := result.rowIndex + 1

		if result.skipped {
			fmt.Printf("Row %d: %s\n", rowNum, result.message)
			continue
		}

		// Write full address
		colName, _ := excelize.ColumnNumberToName(addressCol + 1)
		cell := fmt.Sprintf("%s%d", colName, rowNum)
		s.repo.SetCellValue(cell, result.address)

		// Write district
		colName, _ = excelize.ColumnNumberToName(districtCol + 1)
		cell = fmt.Sprintf("%s%d", colName, rowNum)
		s.repo.SetCellValue(cell, result.district)

		// Write province
		colName, _ = excelize.ColumnNumberToName(provinceCol + 1)
		cell = fmt.Sprintf("%s%d", colName, rowNum)
		s.repo.SetCellValue(cell, result.province)

		fmt.Printf("Row %d: ✓ [%d/%d] (%.6f, %.6f) -> %s\n", rowNum, completed, total, result.coords.Lat, result.coords.Lng, result.address)
		processed++
	}

	return processed
}

// reverseGeocode converts latitude and longitude to full address, district, and province using Nominatim API
func (s *Service) reverseGeocode(lat, lng float64) (address, district, province string, err error) {
	maxRetries := 3
	baseDelay := 2 * time.Second

	for attempt := 0; attempt < maxRetries; attempt++ {
		if attempt > 0 {
			// Exponential backoff: 2s, 4s, 8s
			delay := baseDelay * time.Duration(1<<uint(attempt-1))
			time.Sleep(delay)
		}

		// Using OpenStreetMap Nominatim API (free, no API key required)
		baseURL := "https://nominatim.openstreetmap.org/reverse"

		params := url.Values{}
		params.Set("lat", fmt.Sprintf("%.6f", lat))
		params.Set("lon", fmt.Sprintf("%.6f", lng))
		params.Set("format", "json")
		params.Set("addressdetails", "1")
		params.Set("accept-language", "en") // Request English language

		reqURL := fmt.Sprintf("%s?%s", baseURL, params.Encode())

		// Create HTTP request with proper headers (required by Nominatim)
		req, err := http.NewRequest("GET", reqURL, nil)
		if err != nil {
			return "", "", "", err
		}

		// Better User-Agent identification (required by Nominatim policy)
		req.Header.Set("User-Agent", "latlg-address-converter/1.0")
		req.Header.Set("Accept-Language", "en")
		req.Header.Set("Referer", "https://github.com")

		client := &http.Client{
			Timeout: 15 * time.Second,
		}

		resp, err := client.Do(req)
		if err != nil {
			if attempt < maxRetries-1 {
				continue // Retry on network errors
			}
			return "", "", "", err
		}

		// Handle rate limiting (429) with retry
		if resp.StatusCode == 429 {
			resp.Body.Close()
			if attempt < maxRetries-1 {
				// Wait longer for rate limit
				waitTime := time.Duration(attempt+1) * 10 * time.Second
				time.Sleep(waitTime)
				continue
			}
			return "", "", "", fmt.Errorf("API rate limit exceeded after %d retries", maxRetries)
		}

		if resp.StatusCode != http.StatusOK {
			body, _ := io.ReadAll(resp.Body)
			resp.Body.Close()
			if attempt < maxRetries-1 && resp.StatusCode >= 500 {
				continue // Retry on server errors
			}
			return "", "", "", fmt.Errorf("API returned status %d: %s", resp.StatusCode, string(body))
		}

		var geocodeResp GeocodeResponse
		if err := json.NewDecoder(resp.Body).Decode(&geocodeResp); err != nil {
			resp.Body.Close()
			if attempt < maxRetries-1 {
				continue // Retry on decode errors
			}
			return "", "", "", err
		}
		resp.Body.Close()

		if geocodeResp.DisplayName == "" {
			return "", "", "", fmt.Errorf("no address found for coordinates")
		}

		// Format full address and extract district and province
		address = s.formatFullAddress(geocodeResp)
		district, province = s.extractDistrictAndProvince(geocodeResp)
		return address, district, province, nil
	}

	return "", "", "", fmt.Errorf("failed after %d retries", maxRetries)
}

// formatFullAddress formats the complete address in English
func (s *Service) formatFullAddress(resp GeocodeResponse) string {
	addr := resp.Address
	var parts []string

	// Add road/house number if available
	if addr.HouseNumber != "" && addr.Road != "" {
		parts = append(parts, fmt.Sprintf("%s %s", addr.HouseNumber, addr.Road))
	} else if addr.Road != "" {
		parts = append(parts, addr.Road)
	} else if addr.HouseNumber != "" {
		parts = append(parts, addr.HouseNumber)
	}

	// Add subdistrict/suburb if available
	if addr.Subdistrict != "" {
		parts = append(parts, addr.Subdistrict)
	} else if addr.Suburb != "" {
		parts = append(parts, addr.Suburb)
	}

	// Add district (in English)
	if addr.District != "" {
		parts = append(parts, addr.District)
	} else if addr.County != "" {
		parts = append(parts, addr.County)
	} else if addr.StateDistrict != "" {
		parts = append(parts, addr.StateDistrict)
	}

	// Add province/state (in English)
	if addr.Province != "" {
		parts = append(parts, addr.Province)
	} else if addr.State != "" {
		parts = append(parts, addr.State)
	} else if addr.City != "" {
		parts = append(parts, addr.City)
	}

	// Add postcode if available
	if addr.Postcode != "" {
		parts = append(parts, addr.Postcode)
	}

	// Add country if available
	if addr.Country != "" {
		parts = append(parts, addr.Country)
	}

	// If no parts, return display name as fallback
	if len(parts) == 0 {
		return resp.DisplayName
	}

	return strings.Join(parts, ", ")
}

// extractDistrictAndProvince extracts district and province from the geocode response
func (s *Service) extractDistrictAndProvince(resp GeocodeResponse) (district, province string) {
	addr := resp.Address

	// Extract district (in English) - try multiple fallbacks
	// For Cambodia, district might be in different fields
	if addr.District != "" {
		district = addr.District
	} else if addr.County != "" {
		district = addr.County
	} else if addr.StateDistrict != "" {
		district = addr.StateDistrict
	} else if addr.Subdistrict != "" {
		district = addr.Subdistrict
	} else if addr.Suburb != "" {
		district = addr.Suburb
	} else if addr.City != "" && addr.Province == "" {
		// Use city as district if province is separate
		district = addr.City
	} else {
		// Try to extract from display_name if available
		district = s.extractDistrictFromDisplayName(resp.DisplayName, resp)
	}

	// Extract province/state (in English)
	if addr.Province != "" {
		province = addr.Province
	} else if addr.State != "" {
		province = addr.State
	} else if addr.City != "" {
		province = addr.City
	} else if addr.Country != "" {
		// Fallback to country if province not found
		province = addr.Country
	}

	return district, province
}

// extractDistrictFromDisplayName tries to extract district from the display name
func (s *Service) extractDistrictFromDisplayName(displayName string, addr GeocodeResponse) string {
	// For Cambodia addresses, the structure might be: Road, Subdistrict, District, Province, Country
	// Try to parse common patterns
	parts := strings.Split(displayName, ",")

	// If we have multiple parts, district might be in the middle
	// Common pattern: [road], [subdistrict], [district], [province], [country]
	if len(parts) >= 3 {
		// District is usually the third part from the end (before province and country)
		// Or second part if there's no subdistrict
		for i := len(parts) - 3; i >= 0 && i < len(parts)-1; i-- {
			part := strings.TrimSpace(parts[i])
			// Skip if it's a number (postcode) or known non-district fields
			if part != "" && part != addr.Address.Province &&
				part != addr.Address.Country &&
				part != addr.Address.City && part != addr.Address.State {
				// This might be the district
				return part
			}
		}
	}

	return ""
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run main.go <excel-file.xlsx>")
		fmt.Println("Example: go run main.go coordinates.xlsx")
		fmt.Println("Note: Input file must be in data/ directory, output will be saved to data/")
		os.Exit(1)
	}

	fileName := os.Args[1]

	// Ensure data/ directory exists
	dataDir := "data"
	if err := os.MkdirAll(dataDir, 0755); err != nil {
		log.Fatalf("Error creating data directory: %v", err)
	}

	// Excel file must be in data/ directory
	excelFile := filepath.Join(dataDir, fileName)
	if _, err := os.Stat(excelFile); os.IsNotExist(err) {
		log.Fatalf("Error: File '%s' not found in data/ directory. Please place your Excel file in the data/ folder.", fileName)
	}

	repo, err := NewRepository(excelFile)
	if err != nil {
		log.Fatalf("Error: %v", err)
	}
	defer repo.Close()

	service := NewService(repo)
	if err := service.Process(excelFile); err != nil {
		log.Fatalf("Error: %v", err)
	}
}
