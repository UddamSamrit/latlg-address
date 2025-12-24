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
	repo *Repository
}

// NewService creates a new service instance
func NewService(repo *Repository) *Service {
	return &Service{
		repo: repo,
	}
}

// Process converts coordinates to addresses and saves the result
func (s *Service) Process(excelFile string) error {
	fmt.Printf("Processing sheet: %s\n", s.repo.GetSheetName())

	rows := s.repo.GetRows()
	latLngCol, addressCol, districtCol, provinceCol, err := s.findColumns(rows)
	if err != nil {
		return err
	}

	if addressCol == -1 || districtCol == -1 || provinceCol == -1 {
		addressCol, districtCol, provinceCol = s.addAddressColumns(len(rows[0]))
	}

	processed := s.processRows(rows, latLngCol, addressCol, districtCol, provinceCol)

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

	fmt.Printf("\n✓ Processed %d rows\n", processed)
	fmt.Printf("✓ Output saved to: %s\n", outputFile)
	return nil
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

// processRows processes all data rows and converts coordinates to addresses
func (s *Service) processRows(rows [][]string, latLngCol, addressCol, districtCol, provinceCol int) int {
	processed := 0

	for i := 1; i < len(rows); i++ {
		row := rows[i]

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
			fmt.Printf("Row %d: Skipping empty coordinates\n", i+1)
			continue
		}

		coords, err := s.parseCoordinates(coordStr)
		if err != nil {
			fmt.Printf("Row %d: %v\n", i+1, err)
			continue
		}

		fmt.Printf("Row %d: Processing coordinates (%.6f, %.6f)... ", i+1, coords.Lat, coords.Lng)

		address, district, province, err := s.reverseGeocode(coords.Lat, coords.Lng)
		if err != nil {
			fmt.Printf("Error: %v\n", err)
			continue
		}

		// Write full address
		colName, _ := excelize.ColumnNumberToName(addressCol + 1)
		cell := fmt.Sprintf("%s%d", colName, i+1)
		s.repo.SetCellValue(cell, address)

		// Write district
		colName, _ = excelize.ColumnNumberToName(districtCol + 1)
		cell = fmt.Sprintf("%s%d", colName, i+1)
		s.repo.SetCellValue(cell, district)

		// Write province
		colName, _ = excelize.ColumnNumberToName(provinceCol + 1)
		cell = fmt.Sprintf("%s%d", colName, i+1)
		s.repo.SetCellValue(cell, province)

		fmt.Printf("✓ Address: %s\n", address)
		processed++

		// Be respectful to the API - add delay between requests
		time.Sleep(1 * time.Second)
	}

	return processed
}

// reverseGeocode converts latitude and longitude to full address, district, and province using Nominatim API
func (s *Service) reverseGeocode(lat, lng float64) (address, district, province string, err error) {
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

	req.Header.Set("User-Agent", "latlg-address-converter/1.0")
	req.Header.Set("Accept-Language", "en") // Request English language

	client := &http.Client{
		Timeout: 10 * time.Second,
	}

	resp, err := client.Do(req)
	if err != nil {
		return "", "", "", err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(resp.Body)
		return "", "", "", fmt.Errorf("API returned status %d: %s", resp.StatusCode, string(body))
	}

	var geocodeResp GeocodeResponse
	if err := json.NewDecoder(resp.Body).Decode(&geocodeResp); err != nil {
		return "", "", "", err
	}

	if geocodeResp.DisplayName == "" {
		return "", "", "", fmt.Errorf("no address found for coordinates")
	}

	// Format full address and extract district and province
	address = s.formatFullAddress(geocodeResp)
	district, province = s.extractDistrictAndProvince(geocodeResp)
	return address, district, province, nil
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

	// Extract district (in English)
	if addr.District != "" {
		district = addr.District
	} else if addr.County != "" {
		district = addr.County
	} else if addr.StateDistrict != "" {
		district = addr.StateDistrict
	}
	// If district not found, leave it empty (simple handling)

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
	// If province not found, leave it empty (simple handling)

	return district, province
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
