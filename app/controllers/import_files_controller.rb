class ImportFilesController < ApplicationController

  before_action :get_bank, only: [:import_government_sheet, :programs, :import_freddie_fixed_rate, :import_conforming_fixed_rate, :home_possible, :conforming_arms, :lp_open_acces_arms, :lp_open_access_105, :lp_open_access, :du_refi_plus_arms, :du_refi_plus_fixed_rate_105, :du_refi_plus_fixed_rate, :dream_big, :high_balance_extra, :freddie_arms, :import_jumbo_sheet, :jumbo_series_d,:jumbo_series_f, :jumbo_series_h, :jumbo_series_i, :jumbo_series_jqm]

  require 'roo'
  require 'roo-xls'

  def index
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
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
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
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
              @program.update(term: @term,interest_type: 0,loan_type: 0)
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
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end

              @program.update(interest_points: @block_hash)
            end
          end
        end

        xlsx.sheet(sheet).each_with_index do |sheet_row, index|
          index = index+ 1
          if sheet_row.include?("Loan Level Price Adjustments")
            (index..xlsx.sheet(sheet).last_row).drop(1).each do |row_no|
              adj_row = xlsx.sheet(sheet).row(row_no)
              if xlsx.sheet(sheet).row(row_no).compact.length > 0
                rr = row_no # (r == 8) / (r == 36) / (r == 56)
                max_column_section = adj_row.compact.count - 1
                (0..max_column_section).each do |max_column|
                  cc = 3 + max_column*8 # (3 / 9 / 15)
                  @block_hash = {}
                  key = ''
                  (0..50).each do |max_row|
                    @data = []
                    (0..7).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = xlsx.sheet(sheet).cell(rrr,ccc)
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      elsif (index == 2)
                        @block_hash[key][value.split[0]] = {}
                        # first_row[c_i]
                        # {"Credit Score"=> {0 => 4.0, 580 => 2.00, 600 => 1.250},
                        # {"Credit Score"=>{15=>nil, 30=>"< 580", 45=>nil, 60=>nil, 75=>nil, 90=>4.0, 105=>nil}}
                        # @block_hash[key][(value.include?("<") ? value.split[0] : nil ] = value if !(value.include?("<"))
                      elsif (index == 6)
                        @block_hash[key][value.split[0]] = value
                      end
                      @data << value
                    end

                    if @data.compact.reject { |c| c.blank? }.length == 0
                      break # terminate the loop
                    end
                  end
                  @program.update(interest_points: @block_hash)
                end
              end
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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def jumbo_series_d
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_D")
        sheet_data = xlsx.sheet(sheet)
        (1..22).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 2 + max_column*6 # (2 / 8 / 14)
              @title = sheet_data.cell(r,cc)
                program_heading = @title.split
                @term =  program_heading[3]
                @interest_type = program_heading[5]
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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def jumbo_series_f
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_F")
        sheet_data = xlsx.sheet(sheet)
        (2..36).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8)/ (r == 23)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 6 + max_column*6 # (6 / 12 / 18)
              @title = sheet_data.cell(r,cc)
              program_heading = @title.split
              # @term =  program_heading[3]
              if program_heading[5] == "ARM"
                @interest_type = 2
                if @title.scan(/\d+/)[0] == "5"
                  @interest_subtype = @title.scan(/\d+/)[0]
                elsif @title.scan(/\d+/)[0] == "7"
                  @interest_subtype = @title.scan(/\d+/)[0]
                elsif @title.scan(/\d+/)[0] == "10"
                  @interest_subtype = @title.scan(/\d+/)[0]
                end
              elsif program_heading[5] == "Fixed"
                @interest_type = 0
              end
              @program = @bank.programs.find_or_create_by(title: @title)
              if @interest_subtype.present?
                @program.update(term: @term,interest_type: @interest_type,loan_type: 0,interest_subtype: @interest_subtype )
              else
                @program.update(term: @term,interest_type: @interest_type,loan_type: 0)
              end
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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def jumbo_series_h
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_H")
        sheet_data = xlsx.sheet(sheet)
        (2..86).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("Jumbo Series H Product and Pricing"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8)/ (r == 28)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4 + max_column*6 # (4 / 10 / 16/ 22)
              @title = sheet_data.cell(r,cc)
              if @title.present?
                  program_heading = @title.split
                  if program_heading[5] == "ARM"
                    @term = 0
                    @interest_type = 2
                    if @title.scan(/\d+/)[0] == "5"
                      @interest_subtype = @title.scan(/\d+/)[0]
                    elsif @title.scan(/\d+/)[0] == "7"
                      @interest_subtype = @title.scan(/\d+/)[0]
                    elsif @title.scan(/\d+/)[0] == "10"
                      @interest_subtype = @title.scan(/\d+/)[0]
                    end
                  elsif program_heading[5] == "Fixed"
                    @interest_type = 0
                    @term =  program_heading[3]
                  end
                  if @title.scan(/\w+/).include?("Refinance")
                    @loan_type = 1
                  else
                    @loan_type = 0
                  end

                @program = @bank.programs.find_or_create_by(title: @title)

                if @interest_subtype.present?
                  @program.update(term: @term,interest_type: @interest_type,loan_type: @loan_type ,interest_subtype: @interest_subtype )
                else
                  @program.update(term: @term,interest_type: @interest_type,loan_type: @loan_type)
                end

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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def jumbo_series_i
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_I")
        sheet_data = xlsx.sheet(sheet)
        (2..32).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8)/ (r == 21)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              @title = sheet_data.cell(r,cc)
              if @title.present?
                program_heading = @title.split
                if @title.scan(/\w+/).include?("ARM")
                  @term = 0
                  @interest_type = 2
                  if @title.scan(/\d+/)[0] == "5"
                    @interest_subtype = @title.scan(/\d+/)[0]
                  elsif @title.scan(/\d+/)[0] == "7"
                    @interest_subtype = @title.scan(/\d+/)[0]
                  elsif @title.scan(/\d+/)[0] == "10"
                    @interest_subtype = @title.scan(/\d+/)[0]
                  end
                elsif @title.scan(/\w+/).include?("Fixed")
                  @interest_type = 0
                  @term =  program_heading[3]
                end

                @program = @bank.programs.find_or_create_by(title: @title)
                if @interest_subtype.present?
                  @program.update(term: @term,interest_type: @interest_type,loan_type: 0 ,interest_subtype: @interest_subtype )
                else
                  @program.update(term: @term,interest_type: @interest_type,loan_type: 0)
                end
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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def jumbo_series_jqm
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_JQM")
        sheet_data = xlsx.sheet(sheet)
        (2..60).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8)/ (r == 21)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 6 + max_column*6 # (6 / 12 / 18)
              @title = sheet_data.cell(r,cc)
              if @title.present?
                program_heading = @title.split
                if @title.scan(/\w+/).include?("Fixed")
                  @interest_type = 0
                  # @term =  program_heading[3]
                elsif @title.scan(/\w+/).include?("ARM")
                  @term = 0
                  @interest_type = 2
                  if @title.scan(/\d+/)[0] == "5"
                    @interest_subtype = @title.scan(/\d+/)[0]
                  elsif @title.scan(/\d+/)[0] == "7"
                    @interest_subtype = @title.scan(/\d+/)[0]
                  elsif @title.scan(/\d+/)[0] == "10"
                    @interest_subtype = @title.scan(/\d+/)[0]
                  end
                end
                @program = @bank.programs.find_or_create_by(title: @title)
                if @interest_subtype.present?
                  @program.update(term: @term,interest_type: @interest_type,loan_type: 0 ,interest_subtype: @interest_subtype )
                else
                  @program.update(term: @term,interest_type: @interest_type,loan_type: 0)
                end
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
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break #terminate the loop
                  end
                end
                @program.update(interest_points: @block_hash)
              end
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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
            end
          end
        end

        # @adjustment_hash = {}
        # @rate_adjustment = {}
        # @title_adjustment = {}
        # @adj_data = []
        # (37..62).each do |max_row|
        #   adjustment_row = sheet_data.row(max_row)
        #   @adjustment_data = []
        #   cc = 2
        #   max_column = adjustment_row.count-1
        #   (0..max_column).each do |adj_column|
        #     rr = max_row + 1
        #     cc = adj_column + 2
        #     value = sheet_data.cell(rr,cc)

        #     if adjustment_row.include?("Higher of LTV/CLTV --->")
        #       @adj_data = adjustment_row
        #     end
  
        #     if value.present?
        #       debugger
        #       if value == "LTV Based Adjustments for 20/25/30 Yr Fixed Jumbo Products" || "Rate Adjustments (Increase to rate)" || "Max Price" || "LTV Based Adjustments for 15 Yr Fixed and All ARM Jumbo Products" || "Rate Fall-Out Pricing Special" || "ARM Info"
        #         @main_key = value
        #         @adjustment_hash[@main_key] = {}
        #       elsif cc == 3
        #         @key = value 
        #         @adjustment_hash[@key] = {}
        #       elsif cc > 3 && cc <= 14
        #         @adjustment_hash[@key][@adj_data[adj_column]] = value
        #       elsif cc == 2 && !adjustment_row.include?("FICO")
        #         @key = value 
        #         @adjustment_hash[@key] = {}
        #       end
        #     end
        #     @adjustment_data << value
        #   end

          
        #   # (16..18).each do |adj|
        #   #   rr = max_row 
        #   #   value = sheet_data.cell(rr,adj)
        #   #   if value.present?
        #   #     if adj == 16
        #   #       @new_key = value
        #   #       @rate_adjustment[@new_key] = {}
        #   #     elsif adj > 16 && adj <= 18 
        #   #       @rate_adjustment[@new_key] = value
        #   #     end
        #   #   end
        #     # if value == "20/25/30 Yr Fixed Only"
        #     #   main_key = value
        #     #   @title_adjustment[main_key] = {}
        #     #   debugger
        #     # end
        #   # end
        #   # if @adjustment_data.compact.length == 0
        #   #   break # terminate the loop
        #   # end
        # end
        # debugger
        # @program.update(adjustments: @adjustment_hash)
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
                  rrr = r + max_row
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
              @program.update(interest_points: @block_hash.to_json)
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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
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
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash.to_json)
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

  def import_jumbo_sheet
    @allAdjustments = {}
    titles = get_titles
    table_names = ["State Adjustments", "Max Price"]
    rows_entities = {}
    changed_columns_fields = [">= 800", "< =60"]
    columns_entities = {}#get_columns_entities
    state_ad_columns = {
      heading_pair: {}
    }
    columns = {
      data: {},
      indexs: [12,13,15,16,17]
    }
    row_numbers = [44, 54, 64]
    columns_numbers = [2, 3]
    unwanted_data = "LTV% -->"
    m_key  = nil
    mm_key = nil
    previous_key = nil
    previous_element = nil
    @all_data = {}
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet.eql?("Jumbo Series_D"))
        sheet_data = xlsx.sheet(sheet)
        (41..72).each do |r|
          row    = sheet_data.row(r)
          status = row.compact.count <= 3 || row.compact.include?("FICO/LTV Adjustments - Loan Amount > $1MM")

          if ((row.compact.count > 1) && status) && (!row.compact.include?("California Wholesale Rate Sheet"))
             # r == 43 / 53 / 63
            rr = r + 1 # (r == 44) / (r == 54) / (r == 64)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 2 + max_column*9 # (2 / 11)
              @title = sheet_data.cell(r,cc)
              if(titles.include?(@title))
                program_heading = @title.split
                @term = program_heading[1]
                @interest_type = program_heading[3]
                @program = @bank.programs.find_or_create_by(title: @title)
                @program.update(term: @term,interest_type: 0,loan_type: 0, sheet_name: sheet)
                @block_hash = {}
                key = ''
                unless(table_names.include?(@title))
                  (0..50).each do |max_row|
                    @data = []
                    (0..7).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      if rrr <= 71
                        ccc = cc + c_i
                        value = sheet_data.cell(rrr,ccc)
                        # get all verticle columns in columns data
                        if row_numbers.include?(rr) && value.present? && !value.eql?("n/a") && value.is_a?(String)
                          columns[:data][ccc] = changed_columns_fields.include?(value) ? (value.eql?("< =60") ? "0" : value.tr("^0-9", '')) : value.split("-").first.split(".").first
                          columns_entities = columns[:data]
                          columns_entities[:indexs]  = [12, 13, 15, 16, 17]
                          columns_entities[:numbers] = [4,5,6,7,9]
                        end
                        # get all horizontal columns in state_ad_columns
                        if columns_numbers.include?(ccc) && value.present? && value.is_a?(String)
                          if value != "LTV% -->"
                            state_ad_columns[value] = changed_columns_fields.include?(value) ? value.tr("^0-9", '') : value.split(" -").first
                          end
                        end

                        if (max_row.eql?(0) && c_i.eql?(0))
                          key = find_key(@title) #!@title.eql?("Feature Adjustments") ? sheet_data.cell(rrr + 1,ccc) + "/" + value.split("%")[0] : sheet_data.cell(rrr,ccc)
                          m_key = key
                          # prepare first level hash
                          @block_hash[m_key] = {}
                          @allAdjustments[m_key] = {} if @allAdjustments.empty? 
                        end

                        rows_entities[:main_pair] = state_ad_columns
                        main_key = rows_entities[:main_pair][value]

                        if ccc.eql?(3) && main_key.present? && !@block_hash[m_key].has_key?(main_key)
                          # prepare second level hash
                          mm_key = main_key
                          @block_hash[m_key][main_key] = {}
                        elsif @title.eql?("Feature Adjustments") && ccc.eql?(2) && main_key.present? && !@block_hash[m_key].has_key?(main_key)
                          # prepare second level hash
                          mm_key = main_key
                          @block_hash[m_key][main_key] = {}
                        end

                        if columns_entities[:numbers].include?(ccc)
                          valume_key = columns_entities[ccc]
                          # find third level hash key
                          begin
                            if valume_key.present? && mm_key.present? && @block_hash[m_key].keys.any? && !@block_hash[m_key][mm_key].has_key?(valume_key)
                              # assign third level key value
                              @block_hash[m_key][mm_key][valume_key] = value
                            end
                          rescue Exception => e
                            puts e
                          end

                          if mm_key.eql?("680") && valume_key.eql?("75")
                            if @title.eql?("FICO/LTV Adjustments - Loan Amount ≤ $1MM") or @title.eql?("FICO/LTV Adjustments - Loan Amount > $1MM")
                              indexing = "0" if @title.eql?("FICO/LTV Adjustments - Loan Amount ≤ $1MM")
                              indexing = "1000000" if @title.eql?("FICO/LTV Adjustments - Loan Amount > $1MM")
                              @allAdjustments[@allAdjustments.keys.first][indexing] = {}
                              @allAdjustments[@allAdjustments.keys.first][indexing] = @block_hash[@block_hash.keys.first]
                              # @program.update(adjustments: @block_hash.to_json)
                              make_adjust(@block_hash, @allAdjustments.keys.first, sheet, @program.id)
                              @all_data[@program.title] = @block_hash
                            else
                              # @program.update(adjustments: @block_hash.to_json)
                              make_adjust(@block_hash, @allAdjustments.keys.first, sheet, @program.id)
                              @all_data[@program.title] = @block_hash
                            end
                          else
                            if @title.eql?("FICO/LTV Adjustments - Loan Amount ≤ $1MM") or @title.eql?("FICO/LTV Adjustments - Loan Amount > $1MM")
                              indexing = "0" if @title.eql?("FICO/LTV Adjustments - Loan Amount ≤ $1MM")
                              indexing = "1000000" if @title.eql?("FICO/LTV Adjustments - Loan Amount > $1MM")
                              @allAdjustments[@allAdjustments.keys.first][indexing] = {}
                              @allAdjustments[@allAdjustments.keys.first][indexing] = @block_hash[@block_hash.keys.first]
                              # @program.update(adjustments: @block_hash.to_json)
                              make_adjust(@block_hash, @program.title, sheet, @program.id)
                              @all_data[@program.title] = @block_hash
                            else
                              # @program.update(adjustments: @block_hash.to_json)
                              make_adjust(@block_hash, @program.title, sheet, @program.id)
                              @all_data[@program.title] = @block_hash
                            end
                          end
                        end
                      end
                    end
                  end
                else
                  (0..50).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      state_ad_columns[value] = value if ccc.eql?(11)
                      if columns[:data].values.include?(value)
                        columns[:data][ccc] =  value + (columns[:data].values.count + 1).to_s if rrr.eql?(44) && value.present?
                      else
                        columns[:data][ccc] = value if rrr.eql?(44) && value.present?
                      end
                      if (c_i == 0)
                        key = @block_hash.empty? ? value : state_ad_columns[value]
                        previous_key = key if @block_hash.empty?
                        @block_hash[key] = {} if @block_hash.empty?
                        @block_hash[previous_key][key] = {} if previous_key.present? && key.present? && previous_key != key
                        previous_element = key if previous_key.present? && key.present? && previous_key != key
                        state_ad_columns[:current_element] = key if previous_key.present? && key.present? && previous_key != key
                      elsif ((c_i == 2 || c_i == 5) && value != "State")
                        if @title != "Max Price" && !@block_hash.empty? && !@block_hash[previous_key].empty?
                          previous_element = value
                        end
                      else
                        if rrr < 62 && ccc != 14 && value.present? && !columns[:data].has_value?(value)
                          @block_hash[previous_key][previous_element] = value
                          @all_data[@program.title] = @block_hash
                        end
                      end
                      @data << value
                    end
                    if @data.compact.reject { |c| c.blank? }.length == 0
                      break # terminate the loop
                    elsif @title.eql?("Max Price")
                      begin
                        @hash1 = Hash[*@data.compact] if @data.compact.include?("20/30 Yr Fixed")
                        @block_hash["Max Price"] = @hash1.merge(Hash[*@data.compact]) if @data.compact.include?("15 Yr Fixed")
                        @block_hash.delete("20/30 Yr Fixed") if @data.compact.include?("15 Yr Fixed")
                        # @program.update(adjustments: @block_hash.to_json)
                        @all_data[@program.title] = @block_hash
                      rescue Exception => e
                      end
                    end
                  end
                  # @program.update(adjustments: @block_hash.to_json)
                  make_adjust(@block_hash, @program.title, sheet, @program.id)
                end
              end
            end
          end
        end
      end
    end

    redirect_to programs_import_file_path(@bank)
  end

  private

  def get_bank
    @bank = Bank.find(params[:id])
  end

  def get_titles
    return ["FICO/LTV Adjustments - Loan Amount ≤ $1MM", "State Adjustments", "FICO/LTV Adjustments - Loan Amount > $1MM", "Feature Adjustments", "Max Price"]
  end

  def table_data
    hash_keys = {
      "FICO/LTV Adjustments" => "LoanAmount/FICO/LTV",
      "Feature Adjustments"  => "Feature/LTV",
      "State Adjustments" => "State",
      "Max Price" => "Max Price"
    }

    return hash_keys
  end

  def make_adjust(block_hash, title, sheet_name, program_id)
    adjustment = Adjustment.find_or_create_by(program_title: title)
    adjustment.data = block_hash
    adjustment.program_title = title
    adjustment.sheet_name = sheet_name
    adjustment.program_ids << program_id unless adjustment.program_ids.include?(program_id)
    adjustment.save
  end

  def find_key(title)
    if title.eql?("FICO/LTV Adjustments - Loan Amount ≤ $1MM") or title.eql?("FICO/LTV Adjustments - Loan Amount > $1MM")
      base_key = table_data[@title.split(" -").first]
    else
      base_key = table_data[@title]
    end

    return base_key
  end
end
