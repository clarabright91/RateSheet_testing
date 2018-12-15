class ImportFilesController < ApplicationController
  before_action :get_bank, only: [:import_government_sheet, :programs, :import_freddie_fixed_rate, :import_conforming_fixed_rate, :home_possible, :conforming_arms, :lp_open_acces_arms, :lp_open_access_105, :lp_open_access, :du_refi_plus_arms, :du_refi_plus_fixed_rate_105, :du_refi_plus_fixed_rate, :dream_big, :high_balance_extra, :freddie_arms]

  require 'roo'
  require 'roo-xls'

  def index
    xlsx = Roo::Spreadsheet.open("/home/yuva/Desktop/Projects/Pure-Loan_last/RateSheetExtractor/OB_New_Penn_Financial_Wholesale5806.xls")
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
    xlsx = Roo::Spreadsheet.open("/home/yuva/Desktop/Projects/Pure-Loan_last/RateSheetExtractor/OB_New_Penn_Financial_Wholesale5806.xls")
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

  def import_freddie_fixed_rate
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Freddie Fixed Rate")
        sheet_data = xlsx.sheet(sheet)

        (1..118).each do |r|
          row = sheet_data.row(r)

          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("Freddie Mac 10yr Super Conforming"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)

              #title
              @title = sheet_data.cell(r,cc)

              #term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end

              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end

              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def import_conforming_fixed_rate
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Conforming Fixed Rate")
        sheet_data = xlsx.sheet(sheet)

        (1..118).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("Fannie Mae 10yr High Balance"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)

              #title
              @title = sheet_data.cell(r,cc)

              #term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end

              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end

              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def home_possible
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Home Possible")
        sheet_data = xlsx.sheet(sheet)

        (1..76).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)

              #title
              @title = sheet_data.cell(r,cc)

              #term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end
              
              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
                @interest_subtype = @title.scan(/\d+/)[0].to_i
              end
              
              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def lp_open_acces_arms
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "LP Open Acces ARMs")
        sheet_data = xlsx.sheet(sheet)
        
        (1..35).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              # title
              @title = sheet_data.cell(r,cc)

              # term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end
              if (@term.nil? && @title.include?("ARM"))
                @term = 0
              end

              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                @interest_subtype = @title.scan(/\d+/)[0].to_i              
              end

              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
            
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, interest_subtype: @interest_subtype)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def lp_open_access_105
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "LP Open Access_105")
        sheet_data = xlsx.sheet(sheet)
        
        (1..61).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("LP Open Access 10yr Fixed >125 LTV"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              # title
              @title = sheet_data.cell(r,cc)
              
              # term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end
              if (@term.nil? && @title.include?("ARM"))
                @term = 0
              end

              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                @interest_subtype = @title.scan(/\d+/)[0].to_i              
              end

              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, interest_subtype: @interest_subtype)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def lp_open_access
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "LP Open Access")
        sheet_data = xlsx.sheet(sheet)
        
        (1..61).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("LP Open Access Super Conforming 10 Yr Fixed"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              # title
              @title = sheet_data.cell(r,cc)
              
              # term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end
              if (@term.nil? && @title.include?("ARM"))
                @term = 0
              end

              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                @interest_subtype = @title.scan(/\d+/)[0].to_i              
              end

              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, interest_subtype: @interest_subtype)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def du_refi_plus_arms
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Du Refi Plus ARMs")
        sheet_data = xlsx.sheet(sheet)
        
        (1..35).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              # title
              @title = sheet_data.cell(r,cc)
              
              # term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end
              if (@term.nil? && @title.include?("ARM"))
                @term = 0
              end

              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                @interest_subtype = @title.scan(/\d+/)[0].to_i              
              end

              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, interest_subtype: @interest_subtype)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def du_refi_plus_fixed_rate_105
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Du Refi Plus Fixed Rate_105")
        sheet_data = xlsx.sheet(sheet)
        
        (1..61).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("DU Refi Plus 10yr Fixed >125 LTV"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              # title
              @title = sheet_data.cell(r,cc)
              
              # term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end
              if (@term.nil? && @title.include?("ARM"))
                @term = 0
              end

              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                @interest_subtype = @title.scan(/\d+/)[0].to_i              
              end

              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, interest_subtype: @interest_subtype)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def du_refi_plus_fixed_rate
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Du Refi Plus Fixed Rate")
        sheet_data = xlsx.sheet(sheet)
        
        (1..61).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("DU Refi Plus 10yr Fixed High Balance"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              # title
              @title = sheet_data.cell(r,cc)
              
              # term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end
              if (@term.nil? && @title.include?("ARM"))
                @term = 0
              end

              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                @interest_subtype = @title.scan(/\d+/)[0].to_i              
              end

              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, interest_subtype: @interest_subtype)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def dream_big
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Dream Big")
        sheet_data = xlsx.sheet(sheet)
        
        (1..33).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("Dream Big Jumbo"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 2 + max_column*6 # (3 / 9 / 15)
              # title
              @title = sheet_data.cell(r,cc)
              
              # term
              @term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end
              if (@term.nil? && @title.include?("ARM"))
                @term = 0
              end
              
              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                @interest_subtype = @title.scan(/\d+/)[0].to_i              
              end

              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, interest_subtype: @interest_subtype)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def high_balance_extra
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "High Balance Extra")
        sheet_data = xlsx.sheet(sheet)
        
        (1..23).each do |r|
          row = sheet_data.row(r)
          if (row.compact.include?("High Balance Extra 30 Yr Fixed"))
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 2 + max_column*6 # (3 / 9 / 15)
              # title
              @title = sheet_data.cell(r,cc)
              
              # term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end
              if (@term.nil? && @title.include?("ARM"))
                @term = 0
              end

              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                @interest_subtype = @title.scan(/\d+/)[0].to_i              
              end

              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, interest_subtype: @interest_subtype)
              @block_hash = {}
              key = ''
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = r + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[key] = {}
                  else
                    # first_row[c_i]
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def freddie_arms
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Freddie ARMs")
        sheet_data = xlsx.sheet(sheet)
        
        (1..47).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("Freddie Mac 10-1 ARM (5-2-5) Super Conforming"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              # title
              @title = sheet_data.cell(r,cc)
              
              # term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end
              if (@term.nil? && @title.include?("ARM"))
                @term = 0
              end

              # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # interest sub type
              if @title.include?("3-1 ARM") || @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM")
                @interest_subtype = @title.scan(/\d+/)[0].to_i              
              end

              # conforming
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end

              # freddie_mac
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end

              # fannie_mae
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, interest_subtype: @interest_subtype)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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

  def conforming_arms
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Conforming ARMs")
        sheet_data = xlsx.sheet(sheet)

        (1..47).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)

              @title = sheet_data.cell(r,cc)
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                @term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                @term = @title.scan(/\d+/)[0]
              end
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              if @title.include?("High Balance")
                @jumbo_high_balance = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, jumbo_high_balance: @jumbo_high_balance)
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
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
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


