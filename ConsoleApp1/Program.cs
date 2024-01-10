using GeoUK;
using GeoUK.Coordinates;
using GeoUK.Ellipsoids;
using GeoUK.Projections;
using Microsoft.Extensions.Configuration;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.Data;

public class Program
{
    static async Task Main(string[] args)
    {
        string excelFileName = "output.xlsx";
        //File path
        var binaryPath = Environment.CurrentDirectory;

        string filePath = GetContentFilePath(binaryPath, excelFileName);

        // Ensure that the directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(filePath));

        //Path to store the json file
        var startupPath = binaryPath.Replace("\\bin\\Debug\\net6.0", "");

        // Full path to appsettings.json
        var appSettingsPath = Path.Combine(startupPath, "appsetting.json");

        
        var configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile(appSettingsPath)
            .Build();

        // Load connection string from appsettings.json
        string connectionString = configuration.GetConnectionString("DefaultConnection");
       
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage())
        {
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                //Query to get data from table
                string query = "SELECT * FROM wp_04t3v45eoy_accession";
                using (MySqlDataAdapter dataAdapter = new MySqlDataAdapter(query, connection))
                {
                    DataTable dataTable = new DataTable();
                    dataAdapter.Fill(dataTable);

                    if (dataTable.Rows.Count > 0)
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                        worksheet.Cells.LoadFromDataTable(dataTable, true);

                        // Get the column indices for geo_x_coordinates and geo_y_coordinates
                        int geo_x_coordinatesIndex = dataTable.Columns["geo_x_coordinates"].Ordinal;
                        int geo_y_coordinatesIndex = dataTable.Columns["geo_y_coordinates"].Ordinal;

                        // Iterate through the rows
                        for (int row = 2; row <= dataTable.Rows.Count + 1; row++)
                        {
                            string geo_x_coordinates = dataTable.Rows[row - 2].ItemArray[geo_x_coordinatesIndex].ToString();
                            string geo_y_coordinates = dataTable.Rows[row - 2].ItemArray[geo_y_coordinatesIndex].ToString();

                            //If both are not null
                            if (!string.IsNullOrEmpty(geo_x_coordinates) && !string.IsNullOrEmpty(geo_y_coordinates))
                            {
                                // Convert geo_x_coordinates and geo_y_coordinates to latitude and longitude
                                var result = ConvertEastNorthToLatLong(double.Parse(geo_x_coordinates), double.Parse(geo_y_coordinates));

                                // Update the worksheet with the latitude and longitude values
                                worksheet.Cells[row, geo_x_coordinatesIndex + 1].Value = result.Latitude;
                                worksheet.Cells[row, geo_y_coordinatesIndex + 1].Value = result.Longitude;
                            }
                            //If geo_y_coordinates is null
                            else if (!string.IsNullOrEmpty(geo_x_coordinates) && string.IsNullOrEmpty(geo_y_coordinates))
                            {
                                // Convert easting and northing to latitude and longitude
                                var result = ConvertEastNorthToLatLong(double.Parse(geo_x_coordinates), 0);

                                worksheet.Cells[row, geo_x_coordinatesIndex + 1].Value = result.Latitude;
                                worksheet.Cells[row, geo_y_coordinatesIndex + 1].Value = null;
                            }
                            //If geo_x_coordinates is null
                            else if (string.IsNullOrEmpty(geo_x_coordinates) && !string.IsNullOrEmpty(geo_y_coordinates))
                            {
                                var result = ConvertEastNorthToLatLong(0, double.Parse(geo_y_coordinates));

                                worksheet.Cells[row, geo_x_coordinatesIndex + 1].Value = null;
                                worksheet.Cells[row, geo_y_coordinatesIndex + 1].Value = result.Longitude;
                            }
                            //If both are not null
                            else
                            {
                                worksheet.Cells[row, geo_x_coordinatesIndex + 1].Value = null;
                                worksheet.Cells[row, geo_y_coordinatesIndex + 1].Value = null;
                            }

                        }

                        // Save the Excel file with updated values
                        FileInfo excelFile = new FileInfo(filePath);
                        package.SaveAs(excelFile);

                        Console.WriteLine("Excel file saved to: " + excelFile.FullName);
                    }
                    else
                    {
                        Console.WriteLine("No data retrieved from the database.");
                    }
                }
            }
        }
    }

    static string GetContentFilePath(string binaryPath, string excelFileName)
    {
        var filePath = binaryPath.Replace("\\bin\\Debug\\net6.0", "\\Content");
        return Path.Combine(filePath, excelFileName);
    }

    // Function to convert easting and northing to latitude and longitude
    static LatitudeLongitude ConvertEastNorthToLatLong(double easting, double northing)
    {
        // Convert to Cartesian
        var cartesian = GeoUK.Convert.ToCartesian(new Airy1830(), new BritishNationalGrid(), new EastingNorthing(easting, northing));

        // ETRS89 is effectively WGS84   
        var wgsCartesian = Transform.Osgb36ToEtrs89(cartesian);

        var wgsLatLong = GeoUK.Convert.ToLatitudeLongitude(new Wgs84(), wgsCartesian);

        return wgsLatLong;
    }

}
