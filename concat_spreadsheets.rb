# Author: Neil Thawani
# Date: 2 September 2015
# Description: CSV and XLSX file parser and data combinator. XLS support pending as-needed.
#
# Instructions:
# 1. Put this file in the parent directory of the files/folders that contain the spreadsheets you wish to parse.
# 2. In the Terminal application, navigate to that directory and run the script.
# e.g.:
# cd ~/Documents/Mailing Lists
# ruby concat_spreadsheets.rb

begin
    require 'spreadsheet' # for .xls files
    require 'rubyXL' # for .xlsx files
    require 'csv' # for .csv files
    require 'pry'
rescue LoadError
    puts "Gem files not installed."
    raise
end

OUTPUT_FILE_NAME = "combined_event_attendance.tsv"
array_of_parsed_tsvs = Array.new

### BEGIN METHOD DEFINITIONS
#def parse_xls(spreadsheet_file)
#	puts 'Parsing ' + spreadsheet_file + '...'
#
#	book = Spreadsheet.open(spreadsheet_file)
#    sheet1 = book.worksheet(0) # can use an index or worksheet name
#    sheet1.each do |row|
#        break if row[0].nil? # if first cell empty
#        puts row.join('\t') # looks like it calls "to_s" on each cell's Value
#    end
#end

def parse_xlsx(spreadsheet_file)
	spreadsheet_file_rows = Array.new

	puts 'Parsing ' + spreadsheet_file + '...'

	book = RubyXL::Parser.parse(spreadsheet_file)
	sheet1 = book[0]

	for row_index in 0..sheet1.count-1
		row_as_tsv = String.new

		sheet1[row_index].cells.each do |column|
            row_as_tsv = row_as_tsv + column.value.to_s.strip.gsub(/\r/, " ").gsub(/\n/, " ")
			if !(sheet1[row_index][sheet1[row_index].cells.size - 1].value.to_s == column.value.to_s)
				row_as_tsv = row_as_tsv + "\t"
			end
		end

		spreadsheet_file_rows << row_as_tsv
	end

	return spreadsheet_file_rows
end

def parse_csv(spreadsheet_file)
	spreadsheet_file_rows = Array.new

	puts 'Parsing ' + spreadsheet_file + '...'
    append_cols_once = 0
	CSV.foreach(spreadsheet_file, :headers => true, :encoding => 'ISO-8859-1') do |row|
		row_as_tsv = String.new

		if append_cols_once == 0
            row.each do |column|
                row_as_tsv = row_as_tsv + column[0].to_s
                if !(row[row.length-1] == column[0].to_s)
                    row_as_tsv = row_as_tsv + "\t"
                end
            end

            append_cols_once = 1
        end

        spreadsheet_file_rows << row_as_tsv
        row_as_tsv = ''

		row.each do |column|
			row_as_tsv = row_as_tsv + column[1].to_s
			if !(row[row.length-1] == column[1].to_s)
				row_as_tsv = row_as_tsv + "\t"
			end
		end

		spreadsheet_file_rows << row_as_tsv
	end

	return spreadsheet_file_rows
end

def parse_and_write_normalized_tsvs_to_file(array_of_parsed_tsvs)
    all_files_in_directory_and_subdirectories = Dir.glob("**/*")
    files_to_parse = Array.new

    puts 'Found the files: '
    all_files_in_directory_and_subdirectories.each do |file|
        if file.end_with?('xls') || file.end_with?('xlsx') || file.end_with?('csv')
            files_to_parse << file
            puts file
        end
    end

    puts ''

    for spreadsheet_file in files_to_parse do
        spreadsheet_file_rows = Array.new

        # if spreadsheet_file.end_with?('xls')
        # 	parse_xls(spreadsheet_file)
        # end

        if spreadsheet_file.end_with?('xlsx')
#            unless spreadsheet_file.start_with?('Wine')
                array_of_parsed_tsvs << parse_xlsx(spreadsheet_file)
#            end
        end

        if spreadsheet_file.end_with?('csv')
            array_of_parsed_tsvs << parse_csv(spreadsheet_file)
        end
    end

    count = 0
    parsed_tsv_filenames = Array.new
    array_of_parsed_tsvs.each do |tsv|
        count = count + 1
        file_name = "tsv_to_parse" + count.to_s + ".tsv"
        file = File.open(file_name, "w")
        file.puts tsv
        file.close

        parsed_tsv_filenames << file_name
    end

    return parsed_tsv_filenames
end

def parse_and_combine_tsvs
    # Get input files
    input_files = Dir.glob("tsv_to_parse*.tsv")

    # Collect/combine headers
    all_headers = input_files.reduce([]) do |all_headers, file|
        header_line = File.open(file, &:gets)     # grab first line
        all_headers | CSV.parse_line(header_line, col_sep: "\t") # parse headers and merge with known ones
    end

    quote_chars = %w(" | ~ ^ & * ')

    # Write combined file
    CSV.open(OUTPUT_FILE_NAME, "w", col_sep: "\t") do |out|
        # Write all headers
        out << all_headers

        # Write rows from each file
        input_files.each do |file|
            # TODO: CHANGE ISO-8859-1 to UTF-8 AND VICE-VERSA IF LEAN LAB FILES ARE NOT CLEANED
            CSV.foreach(file, headers: :first_row, encoding: 'UTF-8', col_sep: "\t", quote_char: quote_chars.shift) do |row|
                begin
                    out << all_headers.map { |header| row[header] }
                rescue CSV::MalformedCSVError
                    quote_chars.empty? ? raise : retry
                end
            end
        end
    end
end

def count_and_record_duplicates_from_combined_sheet
    puts ''
    puts 'Counting duplicate entries...'

    first_name_index = 0
    last_name_index = 0
    is_header = true

    name_array = Array.new

    CSV.foreach(OUTPUT_FILE_NAME, encoding: "ISO-8859-1", col_sep: "\t") do |row|
        if is_header
            first_name_index = row.index("First Name")
            last_name_index = row.index("Last Name")
            is_header = false
        else
            first_name = row[first_name_index] || ''
            last_name = row[last_name_index] || ''
            full_name = first_name + " " + last_name
            name_array << full_name unless full_name.strip.empty?
        end
    end

    duplicate_entries = name_array.find_all { |name| name_array.count(name) > 1 }.reject { |element| element.empty? }

    recorded_duplicate_names = Array.new
    duplicate_entries.each do |name|
        recorded_duplicate_names << name + ", " + duplicate_entries.count(name).to_s
    end
    recorded_duplicate_names.uniq!

    recorded_duplicate_names.each do |entry|
        puts entry
    end
end

def delete_parsed_tsvs(parsed_tsv_filenames)
    parsed_tsv_filenames.each do |file|
        File.delete(file)
    end
end
### END METHOD DEFINITIONS

parsed_tsv_filenames = parse_and_write_normalized_tsvs_to_file(array_of_parsed_tsvs)
parse_and_combine_tsvs()
count_and_record_duplicates_from_combined_sheet()
delete_parsed_tsvs(parsed_tsv_filenames)
