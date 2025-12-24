# LatLng to Address Converter

A Go program that reads latitude and longitude coordinates from an Excel file and converts them to district and province addresses using reverse geocoding.

## Features

- Reads Excel files (.xlsx) from `data/` directory
- Automatically detects latitude and longitude columns
- Converts coordinates to district and province using OpenStreetMap Nominatim API
- Writes district and province in English to separate columns
- Creates a new output file with addresses

## Step-by-Step Setup

### Step 1: Clone the Repository

```bash
git clone <repository-url>
cd latlg-address
```

### Step 2: Install Go

#### For macOS:
```bash
# Using Homebrew
brew install go

# Or download from https://golang.org/dl/
```

#### For Linux:
```bash
# Ubuntu/Debian
sudo apt update
sudo apt install golang-go

# Or download from https://golang.org/dl/
```

#### For Windows:
1. Download Go from https://golang.org/dl/
2. Run the installer
3. Follow the installation wizard

#### Verify Installation:
```bash
go version
```
You should see something like: `go version go1.21.x` or higher

### Step 3: Install Dependencies

```bash
go mod download
```

### Step 4: Prepare Your Excel File

1. Create a `data/` directory in the project root (if it doesn't exist):
```bash
mkdir -p data
```

2. Place your Excel file in the `data/` directory

3. Your Excel file should have the following structure:
   - **First row**: Headers
   - **One column**: Contains coordinates in format `lat,lng` (e.g., `13.536964,105.927722`)
   - The column header should contain "latlg", "lat", "coordinate", or "coord"

#### Example Excel Structure:

| LatLng | District | Province |
|--------|----------|----------|
| 13.536964,105.927722 | | |
| 13.7563,100.5018 | | |

### Step 5: Run the Program

```bash
go run main.go your-file.xlsx
```

Or build and run:
```bash
go build -o latlg-address main.go
./latlg-address your-file.xlsx
```

### Step 6: Check Results

The program will:
- Process all rows with coordinates
- Add "District" and "Province" columns if they don't exist
- Fill in district and province names in English
- Save the output to `data/your-file_with_addresses.xlsx`

## Example Output

After running, your Excel file will look like:

| LatLng | District | Province |
|--------|----------|----------|
| 13.536964,105.927722 | Ubon Ratchathani | Ubon Ratchathani |
| 13.7563,100.5018 | Bangkok Noi | Bangkok |

## Project Structure

```
latlg-address/
├── data/                    # Place your Excel files here
│   ├── your-file.xlsx      # Input file
│   └── your-file_with_addresses.xlsx  # Output file
├── main.go                  # Main program
├── go.mod                   # Go dependencies
└── README.md               # This file
```

## Notes

- **Input files must be in `data/` directory**
- The program uses OpenStreetMap Nominatim API, which is free but has rate limits
- There's a 1-second delay between API requests to be respectful to the service
- District and Province columns will be automatically added if they don't exist
- The program processes all rows except the header row
- Addresses are returned in English

## Troubleshooting

### File not found error:
- Make sure your Excel file is in the `data/` directory
- Check that the filename matches exactly (case-sensitive)

### No district/province found:
- The coordinates might be in a location without clear district/province boundaries
- Try checking the coordinates manually on a map

### API errors:
- Check your internet connection
- The API might be temporarily unavailable, try again later

## API Information

This program uses the OpenStreetMap Nominatim reverse geocoding API, which is free and doesn't require an API key. However, please be respectful of their service:
- Don't make too many requests too quickly
- The program includes a 1-second delay between requests
- For high-volume usage, consider using a commercial geocoding service

