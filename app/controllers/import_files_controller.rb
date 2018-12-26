class ImportFilesController < ApplicationController

  before_action :get_bank, only: [:import_government_sheet, :programs, :import_freddie_fixed_rate, :import_conforming_fixed_rate, :home_possible, :conforming_arms, :lp_open_acces_arms, :lp_open_access_105, :lp_open_access, :du_refi_plus_arms, :du_refi_plus_fixed_rate_105, :du_refi_plus_fixed_rate, :dream_big, :high_balance_extra, :freddie_arms, :jumbo_series_d,:jumbo_series_f, :jumbo_series_h, :jumbo_series_i, :jumbo_series_jqm]

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
        @programs_ids = []
        (1..95).each do |r|
          row = sheet_data.row(r)

          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
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

               # interest type
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end

              # streamline
              if @title.include?("FHA") || @title.include?("VA") || @title.include?("USDA")
                @streamline = true  
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @programs_ids << @program.id
              @program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline)
              @block_hash = {}
              key = ''
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end

              @program.update(base_rate: @block_hash)
            end
          end
        end

        #For Adjustments
        xlsx.sheet(sheet).each_with_index do |sheet_row, index|
          index = index+ 1
          if sheet_row.include?("Loan Level Price Adjustments")
            (index..xlsx.sheet(sheet).last_row).each do |adj_row|
              # First Adjustment
              if xlsx.sheet(sheet).row(adj_row).include?("Credit Score")
                begin
                  rr = adj_row
                  cc = 5
                  @credit_hash = {}
                  main_key = "Credit Score"
                  @credit_hash[main_key] = {}
                  @right_adj = {}
                  (0..9).each do |max_row|
                    @data = []
                    rrr = rr + max_row
                    ccc = cc
                    key = xlsx.sheet(sheet).cell(rrr,ccc)
                    if key.present?
                      if (key.include?("<"))
                        key = 0
                      elsif (key.include?("-"))
                        key = key.split("-").first
                      elsif key.include?("≥")
                        key = key.split.last
                      else
                        key
                      end
                      value = xlsx.sheet(sheet).cell(rrr,ccc+4)
                      right_adj_key = xlsx.sheet(sheet).cell(rrr,ccc+7)
                      right_adj_value = xlsx.sheet(sheet).cell(rrr,ccc+13)
                      raise "value is nil at row = #{rrr} and column = #{ccc}" unless value || key
                      @credit_hash[main_key][key] = value
                      @right_adj[right_adj_key] = right_adj_value
                    end

                  end
                  @adjustment_left = Adjustment.create(data: @credit_hash, sheet_name: sheet, program_ids: @programs_ids)
                  @adjustment_right = Adjustment.create(data: @right_adj, sheet_name: sheet, program_ids: @programs_ids)
                rescue => e
                end
              end
              # Second Adjustment
              if xlsx.sheet(sheet).row(adj_row).include?("Loan Size Adjustments")
                begin
                  rr = adj_row
                  cc = 5
                  @loan_size = {}
                  main_key = "Loan Size / Loan Type"
                  @loan_size[main_key] = {}
                  @loan_size[main_key]["Purchase"] = {}
                  @loan_size[main_key]["Refinance"] = {}
                  (0..5).each do |max_row|
                    @data = []
                    rrr = rr + max_row
                    ccc = cc
                    key = xlsx.sheet(sheet).cell(rrr,ccc)
                    if key.present?
                      if (key.include?("<"))
                        key = 0
                      elsif (key.include?("-"))
                        key = key.split("-").first.tr("^0-9", '')
                      else
                        key
                      end
                      value = xlsx.sheet(sheet).cell(rrr,ccc+4)
                      value1 = xlsx.sheet(sheet).cell(rrr,ccc+5)
                      raise "value is nil at row = #{rrr} and column = #{ccc}" unless value || key
                      @loan_size[main_key]["Purchase"][key] = value
                      @loan_size[main_key]["Refinance"][key] = value1
                    end
                    # debugger
                  end
                  @adjustment = Adjustment.create(data: @loan_size, sheet_name: sheet, program_ids: @programs_ids)
                rescue => e
                end
              end
              # Third Adjustment
              if xlsx.sheet(sheet).row(adj_row).include?("Loan Size Adjustments for VA BPC Loans\n(In addition to standard adjustments)")
                begin
                  rr = adj_row
                  cc = 5
                  @loan_size_va_bpc = {}
                  main_key = "Loan Size / Loan Type / VA BPC"
                  @loan_size_va_bpc[main_key] = {}
                  @loan_size_va_bpc[main_key]["Purchase"] = {}
                  @loan_size_va_bpc[main_key]["Refinance"] = {}
                  (0..4).each do |max_row|
                    @data = []
                    rrr = rr + max_row
                    ccc = cc
                    key = xlsx.sheet(sheet).cell(rrr,ccc)
                    if key.present?
                      if (key.include?("<"))
                        key = 0
                      elsif (key.include?("-"))
                        key = key.split("-").first.tr("^0-9", '')
                      elsif (key.include?("≥"))
                        key = key.split.last.tr("^0-9", '')
                      else
                        key
                      end
                      value = xlsx.sheet(sheet).cell(rrr,ccc+4)
                      value1 = xlsx.sheet(sheet).cell(rrr,ccc+5)
                      raise "value is nil at row = #{rrr} and column = #{ccc}" unless value || key
                      @loan_size_va_bpc[main_key]["Purchase"][key] = value
                      @loan_size_va_bpc[main_key]["Refinance"][key] = value1
                    end
                  end
                  @adjustment = Adjustment.create(data: @loan_size_va_bpc, sheet_name: sheet, program_ids: @programs_ids)
                rescue => e
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
              @block_hash.shift
              @program.update(base_rate: @block_hash)
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
              @block_hash.shift
              @program.update(base_rate: @block_hash)
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
              @block_hash.shift
              @program.update(base_rate: @block_hash)
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
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end
        @adjustment_hash = {}
        primary_key = nil
        secondry_key = nil
        misc_adj_key = nil
        (37..71).each do |r|
          row = sheet_data.row(r)
          if row.compact.count >= 1
            (0..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)

              if value.present?
                if value == "Loan Level Price Adjustments: See Adjustment Caps"
                  primary_key = @key = value
                  @adjustment_hash[@key] = {}
                elsif value == "All LP Open Access ARMs"
                  secondry_key = value
                  @adjustment_hash[@key][value] = {}
                end

                if primary_key && secondry_key && cc == 8 && r > 39
                  # @adjustment_hash[primary_key][secondry_key][value] = {}
                  @adjustment_hash[primary_key][secondry_key][all_lp[:rows][r].values.first] = {}
                end

                if r >= 40 && cc >= 11 && cc != 15
                  begin
                    @adjustment_hash[primary_key][secondry_key][all_lp[:rows][r].values.first][all_lp[cc].values.first] = value if all_lp[:rows][r] && all_lp[cc].values
                  rescue Exception => e
                    puts "For row: #{r} and column: #{cc}"
                  end
                end

                # if misc_adj_key && cc == 15
                #   @adjustment_hash[@key][misc_adj_key] = {}
                #   # debugger
                # end

                # if r >= 47 && cc >= 12
                #   debugger
                #   misc_adj_key = @key = value
                #   @adjustment_hash[@key] = {}
                # end

                # if value == "Misc Adjusters"
                #   @key = value
                #   @adjustment_hash[@key] = {}
                #   debugger
                # elsif misc_adj_key && cc == 15
                #   debugger
                #   @adjustment_hash[@key][misc_adj_key] = {}
                # elsif r >= 47 && cc > 12
                #   bb = value
                #   @adjustment_hash[@key][value] = {}
                #   debugger
                # end
              end
              # debugger if r == 49
              # make_adjust(@adjustment_hash, @program.title, sheet, @program.id)
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
        @adjustment_hash = {}
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
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
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end
        (63..86).each do |r|
          row = sheet_data.row(r)
          if row.compact.count >= 1
            (0..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Loan Level Price Adjustments: See Adjustment Caps"
                  primary_key = @key = value
                  @adjustment_hash[@key] = {}
                elsif value == "All Fixed Conforming > 15yr Terms (All Occupancies)"
                  secondry_key = value
                  @adjustment_hash[primary_key][secondry_key] = {}
                elsif r == 65
                  ltv_key = value
                  if ltv_key.include?("<")
                    ltv_key = 0
                  elsif ltv_key.include?("-")
                    ltv_key = ltv_key.split("-").first
                  elsif ltv_key.include?("≥")
                    ltv_key = ltv_key.split.last
                  else
                    ltv_key  
                  end
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end

                if r == 66
                  # debugger
                end
              end
            end
          end
        end

        # (25..44).each do |r|
        #   row = sheet_data.row(r)
        #   if row.compact.count >= 1
        #     (0..9).each do |max_column|
        #       cc = max_column
        #       value = sheet_data.cell(r,cc)
        #       if value.present?
        #         if value == "Pricing Adjustments" || value == "Cashout (adjustments are cumulative)"
        #           primary_key = @key = value
        #           @adjustment_hash[@key] = {}
        #         elsif value == "All High Balance Extra Loans"
        #           secondry_key = value
        #           @adjustment_hash[primary_key][secondry_key] = {}
        #         end

        #         if r == 27 && cc >= 3
        #           begin
        #             @adjustment_hash[primary_key][secondry_key][high_bal_adjustment[cc].values.first] = {}
        #           rescue Exception => e
        #             puts "For row: #{r} and column: #{cc}"
        #           end
        #         end

        #         if r == 34 && cc >= 3
        #           @adjustment_hash[primary_key][high_bal_adjustment[cc].values.first] = {}
        #         end

        #         if r > 27 && r <= 32 && cc >= 3
        #           @adjustment_hash[primary_key][secondry_key][high_bal_adjustment[cc].values.first][high_bal_adjustment[:rows][r].values.first] = value
        #         end

        #         if r >= 34 && r <= 38 && cc >= 3
        #           @adjustment_hash[primary_key][high_bal_adjustment[cc].values.first][high_bal_adjustment[:rows][r].values.first] = value
        #         end
        #       end
        #     end
        #   end
        # end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def jumbo_series_d
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
    @program_arr =[]
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
                @program_arr  << @program.id
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
                    @block_hash[key] = {} if key.present?
                  else
                    # first_row[c_i]
                    begin
                      @block_hash[key][15*c_i] = value if key.present? &&value.present?
                    rescue Exception => e
                    end
                  end
                  @data << value
                end
                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(base_rate: @block_hash.to_json)
            end
          end
        end

        if @program_arr.any?
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
                  # program_heading = @title.split
                  # @term = program_heading[1]
                  # @interest_type = program_heading[3]
                  # @program = @bank.programs.find_or_create_by(title: @title)
                  # @program.update(term: @term,interest_type: 0,loan_type: 0, sheet_name: sheet)
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
                                make_adjust(@block_hash, @title, sheet, @program_arr, true)
                                puts "#{@title} = #{@block_hash}"
                              else
                                make_adjust(@block_hash, @title, sheet, @program_arr, false)
                                # puts "#{@title} = #{@block_hash}"
                              end
                            else
                              if @title.eql?("FICO/LTV Adjustments - Loan Amount ≤ $1MM") or @title.eql?("FICO/LTV Adjustments - Loan Amount > $1MM")
                                indexing = "0" if @title.eql?("FICO/LTV Adjustments - Loan Amount ≤ $1MM")
                                indexing = "1000000" if @title.eql?("FICO/LTV Adjustments - Loan Amount > $1MM")
                                @allAdjustments[@allAdjustments.keys.first][indexing] = {}
                                @allAdjustments[@allAdjustments.keys.first][indexing] = @block_hash[@block_hash.keys.first]
                                make_adjust(@block_hash, @title, sheet, @program_arr, false)
                                # puts "#{@title} = #{@block_hash}"
                              else
                                make_adjust(@block_hash, @title, sheet, @program_arr, false)
                                # puts "#{@title} = #{@block_hash}"
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
                            # @all_data[@program.title] = @block_hash
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
                          make_adjust(@block_hash, @title, sheet, @program_arr) if @data.compact.include?("15 Yr Fixed")
                          puts "#{@title} = #{@block_hash}" if @data.compact.include?("15 Yr Fixed")
                        rescue Exception => e
                        end
                      end
                    end
                    make_adjust(@block_hash, @title, sheet, @program_arr, false)
                    # puts "#{@title} = #{@block_hash}"
                  end
                end
              end
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
              @block_hash.shift
              @program.update(base_rate: @block_hash.to_json)
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
              @block_hash.shift
              @program.update(base_rate: @block_hash.to_json)
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
              @block_hash.shift
              @program.update(base_rate: @block_hash.to_json)
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
                @block_hash.shift
                @program.update(base_rate: @block_hash)
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
              @block_hash.shift
              @program.update(base_rate: @block_hash.to_json)
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
                @block_hash.shift
                @program.update(base_rate: @block_hash)
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
              @block_hash.shift
              @program.update(base_rate: @block_hash.to_json)
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
                @block_hash.shift
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
        @adjustment_hash = {}
        primary_key = ''
        secondry_key = ''
        (25..44).each do |r|
          row = sheet_data.row(r)
          if row.compact.count >= 1
            (0..9).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Pricing Adjustments" || value == "Cashout (adjustments are cumulative)"
                  primary_key = @key = value
                  @adjustment_hash[@key] = {}
                elsif value == "All High Balance Extra Loans"
                  secondry_key = value
                  @adjustment_hash[primary_key][secondry_key] = {}
                end

                if r == 27 && cc >= 3
                  begin
                    @adjustment_hash[primary_key][secondry_key][high_bal_adjustment[cc].values.first] = {}
                  rescue Exception => e
                    puts "For row: #{r} and column: #{cc}"
                  end
                end

                if r == 34 && cc >= 3
                  @adjustment_hash[primary_key][high_bal_adjustment[cc].values.first] = {}
                end

                if r > 27 && r <= 32 && cc >= 3
                  @adjustment_hash[primary_key][secondry_key][high_bal_adjustment[cc].values.first][high_bal_adjustment[:rows][r].values.first] = value
                end

                if r >= 34 && r <= 38 && cc >= 3
                  @adjustment_hash[primary_key][high_bal_adjustment[cc].values.first][high_bal_adjustment[:rows][r].values.first] = value
                end
              end
            end
          end
        end
        make_adjust(@adjustment_hash, @program.title, sheet, @program.id,status)
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
              @block_hash.shift
              @program.update(base_rate: @block_hash.to_json)
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
              @block_hash.shift
              @program.update(base_rate: @block_hash.to_json)
            end
          end
        end
        @adjustment_hash = {}
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        key = ''
        (25..44).each do |r|
          row = sheet_data.row(r)
          if row.compact.count >= 1
            (0..9).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Pricing Adjustments" || value == "Cashout (adjustments are cumulative)"
                  primary_key = @key = value
                  @adjustment_hash[@key] = {}
                elsif value == "All High Balance Extra Loans"
                  secondry_key = value
                  @adjustment_hash[primary_key][secondry_key] = {}
                elsif value == "Subordinate Financing (adjustments are cumulative)"
                  @key = value
                  @adjustment_hash[@key] = {}
                end

                if r == 27 && cc >= 3
                  begin
                    @adjustment_hash[primary_key][secondry_key][high_bal_adjustment[cc].values.first] = {}
                  rescue Exception => e
                    puts "For row: #{r} and column: #{cc}"
                  end
                end

                if r == 34 && cc >= 3
                  @adjustment_hash[primary_key][high_bal_adjustment[cc].values.first] = {}
                end

                if r > 27 && r <= 32 && cc >= 3
                  @adjustment_hash[primary_key][secondry_key][high_bal_adjustment[cc].values.first][high_bal_adjustment[:rows][r].values.first] = value
                end

                if r >= 34 && r <= 38 && cc >= 3
                  @adjustment_hash[primary_key][high_bal_adjustment[cc].values.first][high_bal_adjustment[:rows][r].values.first] = value
                end
                # if (r == 41 && value == "LTV") || (r == 41 && value == "CLTV")
                #   key = value
                #   @adjustment_hash[key] = {}
                # end
                if r >= 40 && r <= 41 && cc == 7
                  if value == "Max Price"
                    key = value
                    @adjustment_hash[key] = {}
                  else
                    @adjustment_hash[key] = value
                  end
                end

                if r > 41 && r <= 44
                  if cc == 2
                    ltv_key = value
                    if ltv_key.include?("<")
                      ltv_key = 0
                    elsif ltv_key.include?("-")
                      ltv_key = ltv_key.split("-")[0]
                    end
                    @adjustment_hash[@key][ltv_key] = {}
                  elsif cc == 3
                    cltv_key = value
                    cltv_key = cltv_key.split.first
                    @adjustment_hash[@key][cltv_key] = {}
                  end
                  if cc > 3 && cc <= 5
                    @adjustment_hash[@key][ltv_key][high_bal_adjustment[:subordinate][cc].values.first] = value
                    @adjustment_hash[@key][cltv_key][high_bal_adjustment[:subordinate][cc].values.first] = value
                  end
                end
              end
            end
          end
        end
        debugger
        make_adjust(@adjustment_hash, @program.title, sheet, @program.id)
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
              @block_hash.shift
              @program.update(base_rate: @block_hash.to_json)
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
              @block_hash.shift
              @program.update(base_rate: @block_hash.to_json)
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

  def get_titles
    return ["FICO/LTV Adjustments - Loan Amount ≤ $1MM", "State Adjustments", "FICO/LTV Adjustments - Loan Amount > $1MM", "Feature Adjustments", "Max Price"]
  end

  def all_lp
    data = {
      11 => {"< 620" => "0"},
      12 => {"620 - 639" => "620"},
      13 => {"640 - 659" => "640"},
      14 => {"660 - 679" => "660"},
      16 => {"680 - 699" => "680"},
      17 => {"700 - 719" => "700"},
      18 => {"720 - 739" => "720"},
      19 => {">= 740" => "740"},
      rows: {
        40 => {"<= 60" => "0"},
        41 => {"60.01 - 70" => "60.01"},
        42 => {"70.01 - 75" => "70.01"},
        43 => {"75.01 - 80" => "75.01"},
        44 => {"80.01 - 85" => "80.01"},
        45 => {"> 85 "=> "85"}
      }
    }

    return data
  end

  def high_bal_adjustment
    data = {
      4 => {"<= 60" => "0"},
      5 => {"60.01 - 70" => "60.01"},
      6 => {"70.01 - 75" => "70.01"},
      7 => {"75.01 - 80" => "75.01"},
      8 => {"80.01 - 85" => "80.01"},
      9 => {"85.01 - 90" => "85"},
      rows: {
        28 => {">=760" => "0"},
        29 => {"740-759" => "740"},
        30 => {"720-739" => "720"},
        31 => {"700-719" => "700"},
        32 => {"680-699" => "680"},
        34 => {">=760" => "0"},
        35 => {"740-759" => "740"},
        36 => {"720-739" => "720"},
        37 => {"700-719" => "700"},
        38 => {"680-699" => "680"}
      },
      subordinate: {
        4 => {"< 720" => "0"},
        5 => {">= 720" => "720"}
      }
    }

    return data
  end

  def high_bal_adjustment
    data = {
      4 => {"<= 60" => "0"},
      5 => {"60.01 - 70" => "60.01"},
      6 => {"70.01 - 75" => "70.01"},
      7 => {"75.01 - 80" => "75.01"},
      8 => {"80.01 - 85" => "80.01"},
      9 => {"85.01 - 90" => "85"},
      rows: {
        28 => {">=760" => "0"},
        29 => {"740-759" => "740"},
        30 => {"720-739" => "720"},
        31 => {"700-719" => "700"},
        32 => {"680-699" => "680"},
        34 => {">=760" => "0"},
        35 => {"740-759" => "740"},
        36 => {"720-739" => "720"},
        37 => {"700-719" => "700"},
        38 => {"680-699" => "680"}
      }
    }

    return data
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
    begin
      adjustment = Adjustment.find_or_create_by(program_title: title)
      adjustment.data = block_hash
      adjustment.sheet_name = sheet_name
      adjustment.program_ids = program_id unless adjustment.program_ids.include?(program_id)
      adjustment.program_ids = adjustment.program_ids.compact.flatten if adjustment.program_ids.include?(nil)
      adjustment.save
    rescue Exception => e
      puts e
    end
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
