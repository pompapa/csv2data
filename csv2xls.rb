require 'csv'
require 'rubyXL'
require 'rubyXL/convenience_methods'

def convert_csv_to_excel(input_file, output_file)
    # Read the CSV file
    csv_data = CSV.read(input_file)

    # Extract the headers from the first row
    headers = csv_data.shift

    # Create a new Excel workbook
    workbook = RubyXL::Workbook.new

    # Add a worksheet to the workbook
    worksheet = workbook[0]

    # Write the headers to the worksheet
    headers.each_with_index do |header, col_index|
        worksheet.add_cell(0, col_index, header)
    end
    
    # Write the data rows to the worksheet
    csv_data.each_with_index do |row, index|
        row.each_with_index do |cell_value, col_index|
            worksheet.add_cell(index + 1, col_index, cell_value)
        end
    end

    # Save the workbook as an Excel file
    workbook.write(output_file)
end

# Specify the input and output file paths from command line arguments
input_file = ARGV[0]
output_file = ARGV[1]

# Convert CSV to Excel
convert_csv_to_excel(input_file, output_file)
# Load the workbook from the output file
workbook = RubyXL::Parser.parse(output_file)

# Access the first worksheet
worksheet = workbook[0]

# Access a specific cell and update its value
## worksheet[0][0].change_contents("New Value")

# Save the updated workbook
workbook.write(output_file)