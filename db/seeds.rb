# This file should ensure the existence of records required to run the application in every environment (production,
# development, test). The code here should be idempotent so that it can be executed at any point in every environment.
# The data can then be loaded with the bin/rails db:seed command (or created alongside the database with db:setup).
#
# Example:
#
#   ["Action", "Comedy", "Drama", "Horror"].each do |genre_name|
#     MovieGenre.find_or_create_by!(name: genre_name)
#   end

# db/seeds.rb
# This seed script reads an Excel file and inserts its data into the database.
# Equivalent to the ExcelJS Node.js script.

# Require the 'roo' gem to work with Excel files.
# Make sure you have it in your Gemfile: gem 'roo'
require 'roo'

begin
  # Load your Excel file
  # Roo automatically detects the file format (.xlsx, .xls, .csv, etc.)
  workbook = Roo::Spreadsheet.open("db/hotels.xlsx")

  # Select the first sheet
  worksheet = workbook.sheet(0)

  # Read headers from the first row
  headers = worksheet.row(1) # => returns an array of header values
  puts "Headers: #{headers.inspect}"

  # Collect rows
  rows = []

  # Iterate over each row, skipping the first one (headers)
  worksheet.each_with_index do |row, index|
    next if index == 0 # skip header row

    # Build a hash mapping header -> cell value
    row_data = {}
    headers.each_with_index do |header, col_index|
      row_data[header] = row[col_index]
    end

    rows << row_data
  end

  puts "Seeding #{rows.size} records..."

  # Example: Insert into ActiveRecord model (replace `Hotel` with your model name)
  rows.each do |row_data|
    # Adjust attributes based on your DB schema
    # Hotel.create!(row_data)
    puts row_data
  end

  puts "✅ Seeding completed!"

rescue => e
  puts "❌ Error while seeding: #{e.message}"
  exit 1
end
