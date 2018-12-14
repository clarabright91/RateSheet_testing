class ImportFilesController < ApplicationController
  before_action :get_bank, only: [:import_government_sheet, :programs, :import_homereddy_sheet, :import_HomeReadyhb_sheet]

  require 'roo'
  require 'roo-xls'

  def index
    xlsx = Roo::Spreadsheet.open("/home/yuva/Desktop/ratesheet/RateSheetExtractor/OB_New_Penn_Financial_Wholesale5806 (1).xls")
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "Cover Zone 1")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          xlsx.sheet(sheet).each_with_index do |row, index|
            current_row = index+1
            if row.include?("Mortgagee Clause (Wholesale)")
              address_index = row.find_index("Mortgagee Clause (Wholesale)")
              @address_a = []
              (1..3).each do |n|
                @address_a << xlsx.sheet(sheet).row(current_row+n)[address_index]
                if n == 3
                  @zip = xlsx.sheet(sheet).row(current_row+n)[address_index].split.last
                  @state_code = xlsx.sheet(sheet).row(current_row+n)[address_index].split[2]
                end
              end
            end
            if (row.include?("Phone") && row.include?("General Contacts"))
              phone_index = row.find_index(headers[0])
              general_contacts_index = row.find_index(headers[1])
              c_row = xlsx.sheet(sheet).row(current_row+1)
              @name = c_row[general_contacts_index]
              @phone = c_row[phone_index]
            end
          end
          @bank = Bank.find_or_create_by(name: @name)
          @bank.update(phone: @phone, address1: @address_a.join, state_code: @state_code, zip: @zip)
        end
      end
    rescue
      # the required headers are not all present
    end
  end

  def import_government_sheet
    xlsx = Roo::Spreadsheet.open("/home/yuva/Desktop/ratesheet/RateSheetExtractor/OB_New_Penn_Financial_Wholesale5806 (1).xls")
    xlsx.sheets.each do |sheet|
      if (sheet == "Government")
        sheet_data = xlsx.sheet(sheet)

        (1..95).each do |r|
          row = sheet_data.row(r)

          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)

              @title = sheet_data.cell(r,cc)
              program_heading = @title.split
              @term = program_heading[1]
              @interest_type = program_heading[3]

              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0)
              @block_hash = {}
              key = ''
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[key] = {}
                  else
                    # first_row[c_i]
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash)
            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end


  def import_homereddy_sheet
    xlsx = Roo::Spreadsheet.open("/home/yuva/Desktop/ratesheet/RateSheetExtractor/OB_New_Penn_Financial_Wholesale5806 (1).xls")
    xlsx.sheets.each do |sheet|
      if (sheet == "HomeReady")
        sheet_data = xlsx.sheet(sheet)

       (1..76).each do |r|
            row = sheet_data.row(r)
            if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
              rr = r + 1 # (r == 8) / (r == 36) / (r == 56)

              max_column_section = row.compact.count - 1
              (0..max_column_section).each do |max_column|

                cc = 3 + max_column*6 # (3 / 9 / 15) 3/8/13

                @title = sheet_data.cell(r,cc)
                program_heading = @title.split


                @term = program_heading[4] == "ARM" ? 0 : program_heading[3]
                @interest_type = program_heading[5] == "Fixed" ? 0 : 2
                if program_heading[3] == "5/1"
                  @interest_subtype = 5
                  elsif program_heading[3] == "7/1"
                    @interest_subtype = 7
                elsif program_heading[3] == "10/1"
                  @interest_subtype = 10
                end


                if @title.include?("Fannie Mae")
                @conforming = true
                @fannie_mae = true
                end

                if @title.include?("Fannie Mae HomeReady")
                  @fannie_mae_home_ready = true
                end

                @program = @bank.programs.find_or_create_by(title: @title)
                @program.update(term: @term,interest_type: @interest_type, interest_subtype: @interest_subtype, loan_type: 0, fannie_mae: @fannie_mae, fannie_mae_home_ready: @fannie_mae_home_ready, conforming: @conforming)
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      # first_row[c_i]
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
              end
              @program.update(interest_points: @block_hash)

            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def import_HomeReadyhb_sheet
    xlsx = Roo::Spreadsheet.open("/home/yuva/Desktop/ratesheet/RateSheetExtractor/OB_New_Penn_Financial_Wholesale5806 (1).xls")
    xlsx.sheets.each do |sheet|
      if (sheet == "HomeReady HB")
        sheet_data = xlsx.sheet(sheet)

       (1..75).each do |r|
            row = sheet_data.row(r)
            if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
              rr = r + 1 # (r == 8) / (r == 36) / (r == 56)

              max_column_section = row.compact.count - 1
              (0..max_column_section).each do |max_column|

                cc = 3 + max_column*6 # (3 / 9 / 15) 3/8/13

                @title = sheet_data.cell(r,cc)
                program_heading = @title.split


                @term = program_heading[4] == "ARM" ? 0 : program_heading[3]
                @interest_type = program_heading[5] == "Fixed" ? 0 : 2
                if program_heading[3] == "5/1"
                  @interest_subtype = 5
                  elsif program_heading[3] == "7/1"
                    @interest_subtype = 7
                elsif program_heading[3] == "10/1"
                  @interest_subtype = 10
                end
                if @title.include?("Fannie Mae")
                @conforming = true
                @fannie_mae = true
                end

                if @title.include?("Fannie Mae HomeReady")
                  @fannie_mae_home_ready = true
                end
                @program = @bank.programs.find_or_create_by(title: @title)
                @program.update(term: @term,interest_type: @interest_type, interest_subtype: @interest_subtype, loan_type: 0, fannie_mae: @fannie_mae, fannie_mae_home_ready: @fannie_mae_home_ready, conforming: @conforming)
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      # first_row[c_i]
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
              end
              @program.update(interest_points: @block_hash)

            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end





  def programs
    @programs = @bank.programs
  end

  private

  def get_bank
    @bank = Bank.find(params[:id])
  end


end


