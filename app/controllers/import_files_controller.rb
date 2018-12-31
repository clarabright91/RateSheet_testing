class ImportFilesController < ApplicationController
  before_action :get_bank, only: [:import_government_sheet, :programs, :import_freddie_fixed_rate, :import_conforming_fixed_rate, :home_possible, :conforming_arms, :lp_open_acces_arms, :lp_open_access_105, :lp_open_access, :du_refi_plus_arms, :du_refi_plus_fixed_rate_105, :du_refi_plus_fixed_rate, :dream_big, :high_balance_extra, :freddie_arms, :jumbo_series_d,:jumbo_series_f, :jumbo_series_h, :jumbo_series_i, :jumbo_series_jqm, :import_homereddy_sheet, :import_HomeReadyhb_sheet]
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
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
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
                      @block_hash[key][15*c_i] = value
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
                  make_adjust(@credit_hash, @programs_ids)
                  make_adjust(@right_adj, @programs_ids)
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
                  end
                  make_adjust(@loan_size, @programs_ids)
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
                  make_adjust(@loan_size_va_bpc, @programs_ids)
                  # @adjustment = Adjustment.create(data: @loan_size_va_bpc, sheet_name: sheet, program_ids: @programs_ids)
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
    program_ids = []
    @allAdjustments = {}
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
              @programs_ids << @program.id
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac)
              @program.adjustments.destroy_all
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

        # adjustments
        previous_title = nil
        @another_title = nil
        modified_keys  = get_table_keys
        data = get_table_keys
        (123..169).each do |r|
          row    = sheet_data.row(r)
          # r == 52 / 68 / 81 / 84 / 89 / 94
          rr = r #+ 1 # (r == 53) / (r == 69) / (r == 82) / (r == 90) / (r == 95)
          max_column_section = row.compact.count - 1
          (0..max_column_section).each do |max_column|
            cc = 3 + max_column * 9 # (2 / 11)
            @title = sheet_data.cell(r,cc)
            @block_hash = {}
            if(@title.eql?("All Fixed Conforming\n(does not apply to terms <=15yrs)"))
              @block_hash[@title] = {}
              key = ''
              another_key = ''
              keyOfHash = ''
              # for Misc Adjusters
              first_key  = ''
              second_key = ''
              third_key  = ''
              (0..49).each do |max_row|
                @data = []
                (3..19).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = index
                  value = sheet_data.cell(rrr,ccc)
                  # implementation of first key
                  if rrr.eql?(132)
                    # for Cash-Out
                    @title = sheet_data.cell(rrr,cc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(138) && index == 3
                    # for Lender Paid MI Adjustments
                    previous_title = @title = sheet_data.cell(rrr,ccc) unless previous_title == @title
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      first_key = "LPMI/PremiumType/FICO"
                      second_key = "LPMI/Term/LTV/FICO"
                      @block_hash[@title][first_key] = {}
                      @block_hash[@title][second_key] = {}
                    end
                  elsif rrr.eql?(155) && index == 3
                    # for Number Of Units
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                    @another_title = "Loan Size Adjustments"
                    @block_hash[@another_title] = {} unless @block_hash.has_key?(@another_title)
                  elsif rrr.eql?(158) && index == 3
                    # for Subordinate Financing
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(163) && index == 3
                    # for Misc Adjusters
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(164) && index.eql?(13)
                    #for Super Conforming Adjustments
                    @another_title = sheet_data.cell(rrr,ccc)
                    @block_hash[@another_title] = {} unless @block_hash.has_key?(@another_title)
                  elsif rrr.eql?(167) && index.eql?(3)
                    #for Non Owner Occupied
                    @another_title = sheet_data.cell(rrr,ccc)
                    @block_hash[@another_title] = {} unless @block_hash.has_key?(@another_title)
                  end

                  #implementation of second key inside first key
                  if rrr > 122 && rrr < 136 && index == 7 && value
                    # for 1st and 2nd table
                    key = get_value(value)
                    @block_hash[@title][key] = {} unless @block_hash[@title].has_key?(key)
                  elsif (rrr > 137) && (rrr < 154)
                    # for Lender Paid MI Adjustments
                    if index == 5 && value
                      key = value
                      if rrr < 144
                        @block_hash[@title][first_key][value] = {} unless @block_hash[@title][first_key].has_key?(value)
                      else
                        @block_hash[@title][second_key][value] = {} unless @block_hash[@title][first_key].has_key?(value)
                      end
                    elsif index == 6 && rrr < 154 && value
                      another_key = get_value(value)
                      @block_hash[@title][second_key][key][another_key] = {} if another_key
                    end
                  end

                  if [156,157].include?(rrr) && ccc == 6
                    # for Number Of Units
                    key = sheet_data.cell(rrr,ccc)
                    @block_hash[@title][key] = {}
                  end

                  if (159..162).to_a.include?(rrr) && ccc < 12
                    # for Subordinate Financing
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      key = get_value(key)
                      @block_hash[@title][key] = {} unless @block_hash[@title].has_key?(key)
                    elsif index.eql?(7)
                      keyOfHash = sheet_data.cell(rrr,ccc)
                      keyOfHash = get_value(keyOfHash)
                      @block_hash[@title][key][keyOfHash] = {}
                    end
                  end

                  if rrr.eql?(155) && [18,19].include?(ccc)
                    # for Loan Size Adjustments
                    another_key = sheet_data.cell(rrr,ccc)
                    another_key = get_value(another_key)
                    @block_hash[@another_title][another_key] = {} unless @block_hash[@another_title].has_key?(another_key)
                  end

                  if (163..166).to_a.include?(rrr) && ccc < 10
                    # for Misc Adjusters
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      @block_hash[@title][key] = {}
                    end
                  end

                  if [167,168,169].include?(rrr) && [7].include?(ccc)
                    #for Non Owner Occupied
                    hash_key = sheet_data.cell(rrr,ccc)
                    hash_key = get_value(hash_key)
                    key = hash_key
                    @block_hash[@another_title][hash_key] = {} if hash_key.present?
                  end

                  if [164,165].include?(rrr) && @another_title
                    # for Super Conforming Adjustments
                    if index.eql?(17)
                      another_key = sheet_data.cell(rrr,ccc)
                      @block_hash[@another_title][another_key] = {} if another_key
                    end
                  end

                  # implementation of third key inside second key with value
                  if rrr > 122 && rrr < 136 && index > 7 && value
                    # for 1st and 2nd table
                    diff_of_row = rrr - 122
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title][key][hash_key] = value unless @block_hash[@title][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 137 && rrr <= 153 && index >= 7 && value
                    # for Lender Paid MI Adjustments
                    diff_of_row = rrr - 137
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = get_value(hash_key)
                    if (138..143).to_a.include?(rrr)
                      @block_hash[@title][first_key][key][hash_key] = value
                    else
                      @block_hash[@title][second_key][key][another_key][hash_key] = value if value
                    end
                  end

                  if [156,157].include?(rrr) && [9,10,11].include?(ccc)
                    # for Number Of Units
                    diff_of_row = rrr - 155
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = get_value(hash_key)
                    @block_hash[@title][key][hash_key] = value if hash_key.present?
                  end

                  if (159..162).to_a.include?(rrr) && ccc > 9 && ccc < 12 && value
                    # for Subordinate Financing
                    diff_of_row = rrr - 158
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = get_value(hash_key)
                    @block_hash[@title][key][keyOfHash][hash_key] = value if hash_key.present?
                  end

                  if (156..163).to_a.include?(rrr) && ccc > 15 && value
                    #for Loan Size Adjustments
                    if ccc.eql?(18)
                      diff_of_column = ccc - 15
                      extra_key = sheet_data.cell(rrr,(ccc-diff_of_column))
                      @block_hash[@another_title]["Purchase"][extra_key] = value
                    else
                      diff_of_column = ccc - 15
                      extra_key = sheet_data.cell(rrr,(ccc-diff_of_column))
                      @block_hash[@another_title]["Refinance"][extra_key] = value
                    end
                  end

                  if (163..166).to_a.include?(rrr) && ccc == 11
                    #for Misc Adjusters
                    first_key = sheet_data.cell(rrr,ccc - 5)
                    @block_hash[@title][first_key] = value
                  end

                  if [167,168,169].include?(rrr) && [11].include?(ccc)
                    #for Non Owner Occupied
                    @block_hash[@another_title][key] = value if key && value
                  end

                  if [164,165].to_a.include?(rrr)
                    # for Super Conforming Adjustments
                    if index.eql?(19)
                      has_key = sheet_data.cell(rrr,ccc)
                      @block_hash[@another_title][another_key][has_key] = value if another_key.present?
                    end
                  end
                end

                @allAdjustments[@title] = @block_hash[@title]
                if @another_title
                  @allAdjustments[@another_title] = @block_hash[@another_title]
                end
              end
            end
          end
        end
      end
    end

    # rename first level keys
    @allAdjustments.keys.each do |key|
      data = get_table_keys
      if data[key]
        @allAdjustments[data[key]] = @allAdjustments.delete(key)
      end
    end

    # create adjustment for each program
    make_adjust(@allAdjustments, program_ids)

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
    @program_ids = []
    @allAdjustments = {}
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Home Possible")
        @sheet = sheet
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
              @program_ids << @program.id
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, sheet_name: sheet)
              @program.adjustments.destroy_all
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

        # adjustments
        previous_title = nil
        @another_title = nil
        modified_keys  = get_table_keys
        data = get_table_keys
        (81..133).each do |r|
          row    = sheet_data.row(r)
          # r == 52 / 68 / 81 / 84 / 89 / 94
          rr = r #+ 1 # (r == 53) / (r == 69) / (r == 82) / (r == 90) / (r == 95)
          max_column_section = 0#row.compact.count - 1
          (0..max_column_section).each do |max_column|
            cc = 3 + max_column * 9 # (2 / 11)
            @title = sheet_data.cell(r,cc)
            @block_hash = {}
            if(@title.eql?("All Conforming\n(does not apply to Fixed terms <=15yrs with LTV <=95)"))
              @block_hash[@title] = {} unless @block_hash.has_key?(@title)
              key = ''
              another_key = ''
              keyOfHash = ''
              # for Misc Adjusters
              first_key  = ''
              second_key = ''
              third_key  = ''
              (0..55).each do |max_row|
                @data = []
                (3..19).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = index
                  value = sheet_data.cell(rrr,ccc)
                  # implementation of first key
                  if rrr.eql?(81)
                    # for All Conforming
                    @title = sheet_data.cell(rrr,cc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(93) && index == 3
                    # for Lender Paid MI Adjustments
                    previous_title = @title = sheet_data.cell(rrr,ccc) unless previous_title == @title
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      first_key = "LPMI/PremiumType/FICO"
                      second_key = "LPMI/Term/LTV/FICO"
                      @block_hash[@title][first_key] = {}
                      @block_hash[@title][second_key] = {}
                    end
                  elsif rrr.eql?(106) && index == 3
                    # for VLIP LPMI Adjustments
                    @title = sheet_data.cell(rrr,cc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(114) && index == 3
                    # for Subordinate Financing
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(114) && index == 13
                    # for Loan Size Adjustments
                    @another_title = sheet_data.cell(rrr,index)
                    @block_hash[@another_title] = {} unless @block_hash.has_key?(@another_title)
                  elsif rrr.eql?(119) && index == 3
                    # for Misc Adjusters
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(122) && index.eql?(3)
                    #for Number Of Units
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(129) && index.eql?(13)
                    # for Adjustment Caps
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  end

                  #implementation of second key inside first key
                  if (81..88).to_a.include?(rrr) && index == 7 && value
                    # for All Conforming
                    key = get_value(value)
                    if key
                      @block_hash[@title][key] = {} unless @block_hash[@title].has_key?(key)
                    end
                  end

                  if (rrr > 92) && (rrr < 105)
                    # for Lender Paid MI Adjustments
                    if index == 5 && value
                      key = value
                      if rrr < 96
                        @block_hash[@title][first_key][value] = {} unless @block_hash[@title][first_key].has_key?(value)
                      else
                        @block_hash[@title][second_key][value] = {} unless @block_hash[@title][first_key].has_key?(value)
                      end
                    elsif index == 6 && rrr > 96 && value
                      another_key = get_value(value)
                      @block_hash[@title][second_key][key][another_key] = {} if another_key
                    end
                  elsif (106..111).to_a.include?(rrr) && index < 7 && value
                    if(rrr == 106) && (ccc == 4)
                      # for VLIP LPMI Adjustments
                      key = sheet_data.cell(rrr,ccc)
                      @block_hash[@title][key] = {}
                    elsif (rrr == 108) && (ccc == 4)
                      first_key  = sheet_data.cell(rrr,ccc)
                      second_key = sheet_data.cell(rrr,ccc + 1)
                      @block_hash[@title][first_key] = {}
                      @block_hash[@title][first_key][second_key] = {} if second_key
                    elsif (rrr > 107) && (ccc == 6)
                      key = get_value(value)
                      @block_hash[@title][first_key][second_key][key] = {} if second_key
                    end
                  end

                  if (114..117).to_a.include?(rrr) && ccc < 10
                    # for Subordinate Financing
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      key = get_value(key)
                      @block_hash[@title][key] = {} unless @block_hash[@title].has_key?(key)
                    elsif index.eql?(7)
                      keyOfHash = sheet_data.cell(rrr,ccc)
                      keyOfHash = get_value(keyOfHash)
                      @block_hash[@title][key][keyOfHash] = {}
                    end
                  end

                  if rrr.eql?(114) && [18,19].include?(ccc)
                    # for Loan Size Adjustments
                    another_key = sheet_data.cell(rrr,ccc)
                    another_key = get_value(another_key)
                    @block_hash[@another_title][another_key] = {} unless @block_hash[@another_title].has_key?(another_key)
                  end

                  if [123,124].include?(rrr) && ccc == 6
                    # for Number Of Units
                    key = sheet_data.cell(rrr,ccc)
                    @block_hash[@title][key] = {}
                  end

                  # implementation of third key inside second key with value
                  if (81..88).to_a.include?(rrr) && index > 9 && value
                    #  for All Conforming
                    diff_of_row = rrr - 80
                    hash_key = sheet_data.cell((rrr -diff_of_row),ccc)
                    hash_key = get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title][key][hash_key] = value unless @block_hash[@title][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 92 && rrr <= 105 && index >= 7 && value
                    # for Lender Paid MI Adjustments
                    diff_of_row = rrr - 92
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = get_value(hash_key)
                    if [93,94,95].include?(rrr)
                      @block_hash[@title][first_key][key][hash_key] = value
                    else
                      @block_hash[@title][second_key][key][another_key][hash_key] = value if value
                    end
                  end

                  if((106..111).to_a.include?(rrr) && (ccc > 6))
                    # for VLIP LPMI Adjustments
                    diff_of_row = rrr - 92
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = get_value(hash_key)
                    if(rrr == 106)
                      # for VLIP LPMI Adjustments
                      @block_hash[@title][key][hash_key] = value if value && hash_key
                    elsif (108..111).to_a.include?(rrr)
                      @block_hash[@title][first_key][second_key][key][hash_key] = value if value && hash_key
                    end
                  elsif (114..117).to_a.include?(rrr) && ccc > 9 && ccc < 12 && value
                    # for Subordinate Financing
                    diff_of_row = rrr - 113
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = get_value(hash_key)
                    @block_hash[@title][key][keyOfHash][hash_key] = value if hash_key.present?
                  end

                  if [119,120].include?(rrr) && ccc == 11
                    #for Misc Adjusters
                    first_key = sheet_data.cell(rrr,ccc - 5)
                    @block_hash[@title][first_key] = value
                  end

                  if (115..122).to_a.include?(rrr) && ccc > 15 && value
                    #for Loan Size Adjustments
                    if ccc.eql?(18)
                      diff_of_column = ccc - 15
                      extra_key = sheet_data.cell(rrr,(ccc-diff_of_column))
                      @block_hash[@another_title]["Purchase"][extra_key] = value
                    else
                      diff_of_column = ccc - 15
                      extra_key = sheet_data.cell(rrr,(ccc-diff_of_column))
                      @block_hash[@another_title]["Refinance"][extra_key] = value
                    end
                  end

                  if [123,124].include?(rrr) && [9,10,11].include?(ccc)
                    # for Number Of Units
                    diff_of_row = rrr - 122
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = get_value(hash_key)
                    @block_hash[@title][key][hash_key] = value if hash_key.present?
                  end
                end

                @allAdjustments[@title] = @block_hash[@title]
                if @another_title
                  @allAdjustments[@another_title] = @block_hash[@another_title]
                end
              end
            end
          end
        end
      end
    end

    # rename first level keys
    @allAdjustments.keys.each do |key|
      data = get_table_keys
      if data[key]
        @allAdjustments[data[key]] = @allAdjustments.delete(key)
      end
    end
    # create adjustment for each program
    make_adjust(@allAdjustments, @program_ids)

    redirect_to programs_import_file_path(@bank, sheet: @sheet)
  end

  def lp_open_acces_arms
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "LP Open Acces ARMs")
        sheet_data = xlsx.sheet(sheet)
        @adjustment_hash = {}
        @programs_ids = []
        primary_key = ''
        secondry_key = ''
        misc_adj_key = ''
        term_key = ''
        ltv_key = ''
        misc_key = ''
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
              @programs_ids << @program.id
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

        (37..71).each do |r|
          row = sheet_data.row(r)
          if row.compact.count >= 1
            (0..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)

              if value.present?
                if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
                  primary_key = value
                  @adjustment_hash[primary_key] = {}
                end
                if value == "All LP Open Access ARMs" || value == "Subordinate Financing" || value == "Number Of Units" || value == "Loan Size Adjustments"
                  secondry_key = value
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                if r == 39 && cc >= 11 && cc <= 19 && cc != 15
                  @adjustment_hash[primary_key][secondry_key][all_lp[cc].values.first] = {}
                end
                if r > 39 && r <= 45 && cc >= 11 && cc <= 19 && cc != 15
                  @adjustment_hash[primary_key][secondry_key][all_lp[cc].values.first][all_lp[:rows][r].values.first] = value
                end
                if r >= 48 && r <= 54 && cc > 2 && cc <= 6
                  unless @adjustment_hash[primary_key][secondry_key].has_key?(all_lp[:rows][r].values.first)
                    @adjustment_hash[primary_key][secondry_key][all_lp[:rows][r].values.first] = {}
                    @adjustment_hash[primary_key][secondry_key][all_lp[:rows][r].values.first][all_lp[:cltv][r].values.first] = {}
                  else
                    @adjustment_hash[primary_key][secondry_key][all_lp[:rows][r].values.first][all_lp[:cltv][r].values.first] = {}
                  end
                end
                if r >= 48 && r <= 54 && cc >= 9 && cc <= 10
                  @adjustment_hash[primary_key][secondry_key][all_lp[:rows][r].values.first][all_lp[:cltv][r].values.first][all_lp[cc].values.first] = value
                end

                if r > 56 && r <= 58 && cc == 3
                  @adjustment_hash[primary_key][secondry_key][all_lp[:rows][r].values.first] = {}
                end
                if r > 56 && r <= 58 && cc > 3 && cc <= 7
                  @adjustment_hash[primary_key][secondry_key][all_lp[:rows][r].values.first][all_lp[cc].values.first] = value
                end
                if r >= 61 && r <= 67 && cc == 6
                  @adjustment_hash[primary_key][secondry_key][all_lp[:rows][r].values.first] = {}
                end
                if r >= 61 && r <= 67 && cc == 10
                  @adjustment_hash[primary_key][secondry_key][all_lp[:rows][r].values.first] = value
                end
                if r >= 69 && r <= 71 && cc == 3
                  secondry_key = value
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                if r >= 69 && r <= 71 && cc == 10
                  @adjustment_hash[primary_key][secondry_key] = value
                end
              end
            end
            (12..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if  value == "Misc Adjusters" || value == "Adjustment Caps"
                  @key = value
                  @adjustment_hash[primary_key][@key] = {}
                end
                if r >= 47 && r <= 57 && cc == 15
                  if value.include?("Condo")
                    misc_key = "Condo=>75.01=>15.01"
                  else
                    misc_key = value
                  end
                  @adjustment_hash[primary_key][@key][misc_key] = {}
                end
                if r >= 47 && r <= 57 && cc == 19
                  @adjustment_hash[primary_key][@key][misc_key] = value
                end

                if r >= 61 && r <= 65 && cc == 16
                  misc_adj_key = value
                  @adjustment_hash[primary_key][@key][misc_adj_key] = {}
                end
                if r > 61 && r <= 65 && cc == 17
                  if value.include?("<")
                    term_key = "0"
                  elsif value.include?(">")
                    term_key = value.split.last
                  else
                    term_key = value
                  end
                  @adjustment_hash[primary_key][@key][misc_adj_key][term_key] = {}
                end
                if r > 61 && r <= 65 && cc == 18
                  if value.include?("<")
                    ltv_key = "0"
                  elsif value.include?(">")
                    ltv_key = value.split.last
                  else
                    ltv_key = value
                  end
                  @adjustment_hash[primary_key][@key][misc_adj_key][term_key][ltv_key] = {}
                end
                if r > 61 && r <= 65 && cc == 19
                  @adjustment_hash[primary_key][@key][misc_adj_key][term_key][ltv_key] = value
                end
                if r >= 67 && r <= 68 && cc == 12
                  misc_adj_key = value
                  @adjustment_hash[primary_key][misc_adj_key] = {}
                end
                if r >= 67 && r <= 68 && cc == 16
                  @adjustment_hash[primary_key][misc_adj_key] = value
                end
              end
            end
          end
        end
        # make_adjust(@adjustment_hash, @program.title, sheet, @program.id)
        Adjustment.create(data: @adjustment_hash,program_title: @program.title, sheet_name: sheet, program_ids: @programs_ids)
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
        @programs_ids = []
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        term_key = ''
        caps_key = ''
        max_key = ''
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
              @programs_ids << @program.id
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
        # Adjustment
        (63..86).each do |r|
          row = sheet_data.row(r)
          if row.compact.count >= 1
            (0..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)

              if value.present?
                if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
                  primary_key = value
                  @adjustment_hash[primary_key] = {}
                end
                if value == "All Fixed Conforming > 15yr Terms (All Occupancies)"
                  secondry_key = "Conforming/RateType/Term/LTV/FICO"
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                if value == "Subordinate Financing"
                  secondry_key = "FinancingType/LTV/CLTV/FICO"
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                if value == "Number Of Units"
                  secondry_key = "PropertyType/LTV"
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                if value == 'Loan Size Adjustments'
                  secondry_key = "Loan Size Adjustments"
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                if r == 66 && cc == 6
                  if value.include?(">")
                    ltv_key = value.split.last
                  end
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r == 66 && cc > 6 && cc <= 19
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = value
                end
                if r == 69 && cc == 5
                  ltv_key = value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r == 69 && cc == 6
                  if value.include?(">")
                    cltv_key = value.split.last
                    @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
                  end
                end
                if r == 69 && cc >= 9 && cc <= 10
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][all_lp[cc].values.first] = {}
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][all_lp[cc].values.first] = value
                end
                if r >= 72 && r <= 73 && cc == 3
                  ltv_key = value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 72 && r <= 73 && cc == 5
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = value
                end
                if r >= 76 && r <= 82 && cc == 6
                  if value.include?("<")
                    ltv_key = "0"
                  elsif value.include?("-")
                    ltv_key = value.split.last
                  end
                end
                if r == 69 && cc >= 9 && cc <= 10
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][all_lp[cc].values.first] = {}
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][all_lp[cc].values.first] = value
                end
                if r >= 72 && r <= 73 && cc == 3
                  ltv_key = value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 72 && r <= 73 && cc == 5
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = value
                end
                if r >= 76 && r <= 82 && cc == 6
                  if value.include?("<")
                    ltv_key = "0"
                  elsif value.include?("-")
                    ltv_key = value.split.last
                  end
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 76 && r <= 82 && cc == 10
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = value
                end
                if r >= 84 && r <= 86 && cc == 3
                  ltv_key = value
                  @adjustment_hash[primary_key][ltv_key] = {}
                end
                if r >= 84 && r <= 86 && cc == 10
                  @adjustment_hash[primary_key][ltv_key] = value
                end
              end
            end
            (12..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if  value == "Misc Adjusters" || value == "Adjustment Caps"
                  @key = value
                  @adjustment_hash[primary_key][@key] = {}
                end
                if r >= 68 && r <= 72 && cc == 15
                  if value.include?("Condo")
                    cltv_key = "Condo=>105=>15.01"
                  else
                    cltv_key = value
                  end
                  @adjustment_hash[primary_key][@key][cltv_key] = {}
                end
                if r >= 68 && r <= 72 && cc == 19
                  @adjustment_hash[primary_key][@key][cltv_key] = value
                end
                if r > 76 && r <= 79 && cc == 16
                  caps_key = value
                  @adjustment_hash[primary_key][@key][caps_key] = {}
                end
                if r > 76 && r <= 79 && cc == 17
                  if value.include?("<")
                    term_key = "0"
                  elsif value.include?(">")
                    term_key = value.split.last
                  else
                    term_key = value
                  end
                  @adjustment_hash[primary_key][@key][caps_key][term_key] = {}
                end
                if r > 76 && r <= 79 && cc == 18
                  if value.include?("<")
                    ltv_key = "0"
                  elsif value.include?(">")
                    ltv_key = value.split.last
                  else
                    ltv_key = value
                  end
                  @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = {}
                end
                if r > 76 && r <= 79 && cc == 19
                  @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = value
                end
                if r == 82 && cc == 12
                  max_key = value
                  @adjustment_hash[primary_key][max_key] = {}
                end
                if r == 82 && cc == 16
                  @adjustment_hash[primary_key][max_key] = value
                end
              end
            end
          end
        end
        # make_adjust(@adjustment_hash, @program.title, sheet, @programs_ids)
        Adjustment.create(data: @adjustment_hash,program_title: @program.title, sheet_name: sheet, program_ids: @programs_ids)
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
                                make_adjust(@block_hash, @program_arr)
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
                    make_adjust(@block_hash, @program_arr)
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
        @adjustment_hash = {}
        @programs_ids = []
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        unit_key = ''
        caps_key = ''
        term_key = ''
        max_key = ''
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
              @programs_ids << @program.id
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

        # Adjustment
        (63..97).each do |r|
          row = sheet_data.row(r)
          if row.compact.count >= 1
            (0..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)

              if value.present?
                if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
                  primary_key = value
                  @adjustment_hash[primary_key] = {}
                end
                if value == "All Fixed Conforming > 15yr Terms (All Occupancies)"
                  secondry_key = "Conforming/RateType/Term/LTV/FICO"
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                if value == "Subordinate Financing"
                  secondry_key = "FinancingType/LTV/CLTV/FICO"
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                if value == "Number Of Units"
                  secondry_key = "PropertyType/LTV"
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                if value == 'Loan Size Adjustments'
                  secondry_key = "Loan Size Adjustments"
                  @adjustment_hash[primary_key][secondry_key] = {}
                end

                # All fixed Adjustment
                if r >= 66 && r <= 71 && cc == 8
                  if value.include?("<")
                    ltv_key = "0"
                  elsif value.include?("-") || value.include?(">")
                    ltv_key = value.split.last
                  end
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 66 && r <= 71 && cc > 8 && cc <= 19 && cc != 15
                  @adjustment_hash[primary_key][secondry_key][ltv_key][all_lp[cc].values.first] = {}
                  @adjustment_hash[primary_key][secondry_key][ltv_key][all_lp[cc].values.first] = value
                end

                # Subordinate Adjustment
                if r >= 74 && r <= 80 && cc == 5
                  if value.include?("<")
                    ltv_key = "0"
                  elsif value.include?(">") || value.include?("-")
                    ltv_key = value.split.last
                  else
                    ltv_key = value
                  end
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 74 && r <= 80 && cc == 6
                  if value.include?("<")
                    cltv_key = "0"
                  elsif value.include?(">") || value.include?("-")
                    cltv_key = value.split.last
                  else
                    cltv_key = value
                  end
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
                end
                if r >= 74 && r <= 80 && cc >= 9 && cc <= 10
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][all_lp[cc].values.first] = {}
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][all_lp[cc].values.first] = value
                end

                # Number of unit Adjustment
                if r >= 83 && r <= 84 && cc == 3
                  unit_key = value
                  @adjustment_hash[primary_key][secondry_key][unit_key] = {}
                end
                if r >= 83 && r <= 84 && cc > 3 && cc <= 7
                  @adjustment_hash[primary_key][secondry_key][unit_key][all_lp[cc].values.first] = {}
                  @adjustment_hash[primary_key][secondry_key][unit_key][all_lp[cc].values.first] = value
                end

                # Loan Size Adjustments
                if r >= 87 && r <= 93 && cc == 6
                  if value.include?("<")
                    ltv_key = "0"
                  elsif value.include?("-")
                    ltv_key = value.split.last
                  end
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 87 && r <= 93 && cc == 10
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = value
                end

                # Other Adjustment
                if r >= 95 && r <= 97 && cc == 3
                  ltv_key = value
                  @adjustment_hash[primary_key][ltv_key] = {}
                end
                if r >= 95 && r <= 97 && cc == 10
                  @adjustment_hash[primary_key][ltv_key] = value
                end
              end
            end
            (12..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if  value == "Misc Adjusters" || value == "Adjustment Caps"
                  @key = value
                  @adjustment_hash[primary_key][@key] = {}
                end

                # Misc Adjustment
                if r >= 73 && r <= 80 && cc == 15
                  if value.include?("Condo")
                    cltv_key = "Condo=>75.01=>15.01"
                  else
                    cltv_key = value
                  end
                  @adjustment_hash[primary_key][@key][cltv_key] = {}
                end
                if r >= 73 && r <= 80 && cc == 19
                  @adjustment_hash[primary_key][@key][cltv_key] = value
                end

                # Adjustment Caps
                if r > 86 && r <= 90 && cc == 16
                  caps_key = value
                  @adjustment_hash[primary_key][@key][caps_key] = {}
                end
                if r > 86 && r <= 90 && cc == 17
                  if value.include?("<")
                    term_key = "0"
                  elsif value.include?(">")
                    term_key = value.split.last
                  else
                    term_key = value
                  end
                  @adjustment_hash[primary_key][@key][caps_key][term_key] = {}
                end
                if r > 86 && r <= 90 && cc == 18
                  if value.include?("<")
                    ltv_key = "0"
                  elsif value.include?(">")
                    ltv_key = value.split.last
                  else
                    ltv_key = value
                  end
                  @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = {}
                end
                if r > 86 && r <= 90 && cc == 19
                  @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = value
                end


                if r == 93 && cc == 12
                  max_key = value
                  @adjustment_hash[primary_key][max_key] = {}
                end
                if r == 93 && cc == 16
                  @adjustment_hash[primary_key][max_key] = value
                end
              end
            end
          end
        end
        Adjustment.create(data: @adjustment_hash,program_title: @program.title, sheet_name: sheet, program_ids: @programs_ids)
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
        (59..sheet_data.last_row).each_with_index do |adj_row, index|
          @hash = {}
          index = index +1
          @hash["Purchase Transactions"] = {}
          @hash["R/T Refinance Transactions"] = {}
          @hash["C/O Refinance Transactions"] = {}
          @hash["State Adjustments"] = {}
          @hash["Max Price"] = {}
          @hash["Loan Amount Adjustments"] = {}
          @hash["Feature Adjustments"] = {}
          @hash["Product Adjustments"] = {}
          @hash["Special Adjustments (Amort ≥ 240 Months - Fixed Products Only)"] = {}
          (adj_row+2..adj_row+6).each do |max_row|

            key_val = ''
            key_val1 = ''
            (3..13).each do |max_column|
              header_r = (adj_row+2) - index
              ccc = max_column
              rrr = max_row
              value = xlsx.sheet(sheet).cell(rrr,ccc)
              value1 = xlsx.sheet(sheet).cell(header_r,ccc)
              if value1.present? 
                if (value1.include?("≤"))
                  value1 = 0
                elsif (value1.include?("-"))
                  value1 = value1.split("-").first
                elsif (value1.include?("≥"))
                  value1 = value1.split("≥").last
                else
                  value1
                end
              end
              if value.present?
                if ccc == 3
                  if (value.include?("≥"))
                    key_val = (value.split("≥").last)
                  elsif (value.include?("-"))
                    key_val = value.split("-").first
                  else
                    key_val = value
                  end
                  @hash["Purchase Transactions"][key_val] = {}
                else
                  @hash["Purchase Transactions"][key_val][value1] = value
                end
              end
            end
            (15..25).each do |max_column|
              header_r = adj_row
              ccc = max_column
              rrr = max_row - 1
              value = xlsx.sheet(sheet).cell(rrr,ccc)
              value1 = xlsx.sheet(sheet).cell(header_r,ccc)
              if value1.present? && value1.class == String
                if (value1.include?("≤"))
                  value1 = 0
                elsif (value1.include?("-"))
                  value1 = value1.split("-").first
                elsif (value1.include?("≥"))
                  value1 = value1.split("≥").last
                else
                  value1
                end
              end
              if value.present?
                if ccc == 15
                  if (value.include?("≤"))
                    key_val1 = 0
                  elsif (value.include?("≥"))
                    key_val1 = (value.split("≥").last)
                  elsif (value.include?("-"))
                    key_val1 = value.split("-").first
                  else
                    key_val1 = value
                  end
                  @hash["Loan Amount Adjustments"][key_val1] = {}
                else
                  @hash["Loan Amount Adjustments"][key_val1][value1] = value
                end
              end
            end
          end

          (adj_row+10..adj_row+15).each do |max_row|
            
            key_val = ''
            key_val1 = ''
            (3..13).each do |max_column|
              header_r = (adj_row+10) - index
              ccc = max_column
              rrr = max_row
              value = xlsx.sheet(sheet).cell(rrr,ccc)
              value1 = xlsx.sheet(sheet).cell(header_r,ccc)
              if value1.present?
                if (value1.include?("≤"))
                  value1 = 0
                elsif (value1.include?("-"))
                  value1 = value1.split("-").first
                elsif (value1.include?("≥"))
                  value1 = value1.split("≥").last
                else
                  value1
                end
              end
              if value.present?
                if ccc == 3
                  if (value.include?("≥"))
                    key_val = (value.split("≥").last)
                  elsif (value.include?("-"))
                    key_val = value.split("-").first
                  else
                    key_val = value
                  end
                  @hash["R/T Refinance Transactions"][key_val] = {}
                else
                  @hash["R/T Refinance Transactions"][key_val][value1] = value if key_val.present? && key_val != ""
                end
              end
            end
            (15..25).each do |max_column|
              header_r = adj_row+8
              ccc = max_column
              rrr = max_row - 1
              value = xlsx.sheet(sheet).cell(rrr,ccc)
              value1 = xlsx.sheet(sheet).cell(header_r,ccc)
              if value1.present? && value1.class == String
                if (value1.include?("≤"))
                  value1 = 0
                elsif (value1.include?("-"))
                  value1 = value1.split("-").first
                elsif (value1.include?("≥"))
                  value1 = value1.split("≥").last
                else
                  value1
                end
              end
              if value.present?
                if ccc == 15
                  key_val1 = value
                  @hash["Loan Amount Adjustments"][key_val1] = {}
                else
                  @hash["Loan Amount Adjustments"][key_val1][value1] = value if key_val1.present?
                end
              end
            end
          end

          (adj_row+18..adj_row+22).each do |max_row|
                        
            key_val = ''
            key_val1 = ''
            (3..13).each do |max_column|
              header_r = (adj_row+18) - index
              ccc = max_column
              rrr = max_row
              value = xlsx.sheet(sheet).cell(rrr,ccc)
              value1 = xlsx.sheet(sheet).cell(header_r,ccc)
              if value1.present?
                if (value1.include?("≤"))
                  value1 = 0
                elsif (value1.include?("-"))
                  value1 = value1.split("-").first
                elsif (value1.include?("≥"))
                  value1 = value1.split("≥").last
                else
                  value1
                end
              end
              if value.present?
                if ccc == 3
                  if (value.include?("≥"))
                    key_val = (value.split("≥").last)
                  elsif (value.include?("-"))
                    key_val = value.split("-").first
                  else
                    key_val = value
                  end
                  @hash["C/O Refinance Transactions"][key_val] = {}
                else
                  @hash["C/O Refinance Transactions"][key_val][value1] = value if key_val.present? && key_val != ""
                end
              end
            end
            (15..25).each do |max_column|
              header_r = adj_row+9
              ccc = max_column
              rrr = max_row + 1
              value = xlsx.sheet(sheet).cell(rrr,ccc)
              value1 = xlsx.sheet(sheet).cell(header_r,ccc)
              if value1.present? && value1.class == String
                if (value1.include?("≤"))
                  value1 = 0
                elsif (value1.include?("-"))
                  value1 = value1.split("-").first
                elsif (value1.include?("≥"))
                  value1 = value1.split("≥").last
                else
                  value1
                end
              end
              if value.present?
                if ccc == 15
                  key_val1 = value
                  @hash["Product Adjustments"][key_val1] = {}
                else
                  @hash["Product Adjustments"][key_val1][value1] = value if key_val1.present?
                end
              end
            end
          end

          (adj_row+27..adj_row+28).each do |max_row|
                        
            key_val = ''
            (3..13).each do |max_column|
              header_r = (adj_row+27) - index
              ccc = max_column
              rrr = max_row
              value = xlsx.sheet(sheet).cell(rrr,ccc)
              value1 = xlsx.sheet(sheet).cell(header_r,ccc)
              if value1.present?
                if (value1.include?("≤"))
                  value1 = 0
                elsif (value1.include?("-"))
                  value1 = value1.split("-").first
                elsif (value1.include?("≥"))
                  value1 = value1.split("≥").last
                else
                  value1
                end
              end
              if value.present?
                if ccc == 3
                  if (value.include?("≥"))
                    key_val = (value.split("≥").last)
                  elsif (value.include?("-"))
                    key_val = value.split("-").first
                  else
                    key_val = value
                  end
                  @hash["State Adjustments"][key_val] = {}
                else
                  @hash["State Adjustments"][key_val][value1] = value if key_val.present? && key_val != ""
                end
              end
            end
          end

          (adj_row+33..adj_row+36).each do |max_row|
                        
            key_val = ''
            (3..13).each do |max_column|
              header_r = (adj_row+33) - index
              ccc = max_column
              rrr = max_row
              value = xlsx.sheet(sheet).cell(rrr,ccc)
              value1 = xlsx.sheet(sheet).cell(header_r,ccc)
              # if value1.present?
              #   if (value1.include?("≤"))
              #     value1 = 0
              #   elsif (value1.include?("-"))
              #     value1 = value1.split("-").first
              #   elsif (value1.include?("≥"))
              #     value1 = value1.split("≥").last
              #   else
              #     value1
              #   end
              # end
              if value.present?
                if ccc == 3
                  if (value.include?("≤"))
                    key_val = 0
                  elsif (value.include?("≥"))
                    key_val = (value.split("≥").last)
                  elsif (value.include?("-"))
                    key_val = value.split("-").first
                  else
                    key_val = value
                  end
                  @hash["Max Price"][key_val] = {}
                else
                  @hash["Max Price"][key_val][value1] = value if value1.present?
                end
              end
            end
          end

          (adj_row+29..adj_row+31).each do |max_row|
                        
            key_val = ''
            (15..25).each do |max_column|
              header_r = (adj_row+33) - index
              ccc = max_column
              rrr = max_row
              value = xlsx.sheet(sheet).cell(rrr,ccc)
              value1 = xlsx.sheet(sheet).cell(header_r,ccc)
              if value1.present?
                if (value1.include?("≤"))
                  value1 = 0
                elsif (value1.include?("-"))
                  value1 = value1.split("-").first
                elsif (value1.include?("≥"))
                  value1 = value1.split("≥").last
                else
                  value1
                end
              end
              if value.present?
                if ccc == 15
                  if (value.include?("≤"))
                    key_val = 0
                  elsif (value.include?("≥"))
                    key_val = (value.split("≥").last)
                  elsif (value.include?("-"))
                    key_val = value.split("-").first
                  else
                    key_val = value
                  end
                  @hash["Special Adjustments (Amort ≥ 240 Months - Fixed Products Only)"][key_val] = {}
                else
                  @hash["Special Adjustments (Amort ≥ 240 Months - Fixed Products Only)"][key_val][value1] = value if value1.present?
                end
              end
            end
          end
          Adjustment.create(data: @hash,program_title: @program.title, sheet_name: sheet, program_ids: @programs_ids)
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
        # Adjustment

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
        @adjustment_hash = {}
        @programs_ids = []
        primary_key = ''
        main_key = ''
        # programs
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
                @programs_ids << @program.id
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
        # Adjustments
        (34..73).each do |r|
          row = sheet_data.row(r)
          (0..10).each do |max_column|
            cc = 3 + max_column
            value = sheet_data.cell(r,cc)
            if value.present?
              if value == "Jumbo Series I Adjustments"
                main_key = value
                @adjustment_hash[main_key] = {}
              end
              if value == "Fixed Adjustments"
                @key1 = value
                @adjustment_hash[main_key][@key1] = {}
              end
              if value == "Credit Score" || value == "Loan Amount" || value == "Purpose/Property Type"
                primary_key = value
                @adjustment_hash[main_key][@key1][primary_key] = {}
              end
              if (r == 40 && cc > 3 && cc <= 10) || (r == 49 && cc > 3 && cc <= 10) || (r == 57 && cc > 3 && cc <= 10)
                @adjustment_hash[main_key][@key1][primary_key][jumbo_series_i_adjustment[cc].values.first] = {}
              end
              if (r > 40 && r <= 45 && cc > 3 && cc <= 10) || (r > 49 && r <= 53 && cc > 3 && cc <= 10) || (r > 57 && r <= 62 && cc > 3 && cc <= 10)
                @adjustment_hash[main_key][@key1][primary_key][jumbo_series_i_adjustment[cc].values.first][jumbo_series_i_adjustment[:rows][r].values.first] = value
              end
            end
          end

          (12..19).each do |max_column|
            cc = max_column
            value = sheet_data.cell(r,cc)
            if value.present?
              if value == "ARM Adjustments"
                @key = value
                @adjustment_hash[main_key][@key] = {} unless @adjustment_hash[main_key].has_key?(@key)
              end
              if value == "Credit Score" || value == "Loan Amount" || value == "Purpose/Property Type"
                primary_key = value
                @adjustment_hash[main_key][@key][primary_key] = {}
              end
              if (r == 40 && cc > 12 && cc <= 19) || (r == 49 && cc > 12 && cc <= 19) || (r == 57 && cc > 12 && cc <= 19)
                @adjustment_hash[main_key][@key][primary_key][jumbo_series_i_adjustment[cc].values.first] = {}
              end
              if (r > 40 && r <= 45 && cc > 12 && cc <= 19) || (r > 49 && r <= 53 && cc > 12 && cc <= 19) || (r > 57 && r <= 62 && cc > 12 && cc <= 19)
                @adjustment_hash[main_key][@key][primary_key][jumbo_series_i_adjustment[cc].values.first][jumbo_series_i_adjustment[:rows][r].values.first] = value
              end
            end
          end

          (0..17).each do |max_column|
            cc = 3 + max_column
            value = sheet_data.cell(r,cc)
            if value.present?
              if value == "Other Adjustments" || value == "Maximum Prices"
                main_key = value
                @adjustment_hash[main_key] = {}
              end
              if (r == 66 && cc == 3) || (r == 66 && cc == 12) || (r >= 72 && r <= 73 && cc == 3) || (r >= 72 && r <= 73 && cc == 12)
                primary_key = value
                @adjustment_hash[main_key][primary_key] = {}
              end
              if (r == 66 && cc == 7) || (r == 66 && cc == 17) || (r >= 72 && r <= 73 && cc ==7) || (r >= 72 && r <= 73 && cc ==17)
                @adjustment_hash[main_key][primary_key] = value
              end
            end
          end
        end
        # make_adjust(@adjustment_hash, @program.title, sheet, @program_ids)
        Adjustment.create(data: @adjustment_hash,program_title: @program.title, sheet_name: sheet, program_ids: @programs_ids)
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
        cltv_key = ''
        cltv_data = ''
        adj_key = ''
        @cltv_data = []
        @cltv_data2 = []
        @max_price_data = []
        (64..99).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(67)
          @cltv_data2 = sheet_data.row(66)
          @max_price_data = sheet_data.row(94)
          if row.compact.count >= 1
            (3..13).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "LTV/FICO Adjustments" 
                  primary_key = value
                  @adjustment_hash[primary_key] = {}
                end
                if value == "Purchase Transactions" || value == "R/T Refinance Transactions" || value == "C/O Refinance Transactions" || value == "State Adjustments"
                  secondry_key = value                  
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                # Purchase Transactions Adjustment
                if r >= 68 && r <= 74 && cc == 3
                  if value.include?("-")
                    cltv_key = value.split("-").last
                  elsif value.include?("≥")
                    cltv_key = value.split("≥").last
                  end
                  @adjustment_hash[primary_key][secondry_key][cltv_key] = {}
                end
                if r >= 68 && r <= 74 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    cltv_data = 0
                  elsif @cltv_data[cc-2].include?("-")
                    cltv_data = @cltv_data[cc-2].split("-").last
                  elsif @cltv_data[cc-2].include?("≥")
                    cltv_data = @cltv_data[cc-2].split("≥").last
                  end
                  @adjustment_hash[primary_key][secondry_key][cltv_key][cltv_data] = {}
                  @adjustment_hash[primary_key][secondry_key][cltv_key][cltv_data] = value
                end

                # R/T Refinance Transactions Adjustment
                if r >= 78 && r <= 84 && cc == 3
                  if value.include?("-")
                    cltv_key = value.split("-").last
                  elsif value.include?("≥")
                    cltv_key = value.split("-").last
                  end
                  @adjustment_hash[primary_key][secondry_key][cltv_key] = {}
                end
                if r >= 78 && r <= 84 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    cltv_data = 0
                  elsif @cltv_data[cc-2].include?("-")
                    cltv_data = @cltv_data[cc-2].split("-").last
                  elsif @cltv_data[cc-2].include?("≥")
                    cltv_data = @cltv_data[cc-2].split("≥").last
                  end
                  @adjustment_hash[primary_key][secondry_key][cltv_key][cltv_data] = {}
                  @adjustment_hash[primary_key][secondry_key][cltv_key][cltv_data] = value
                end

                # C/O Refinance Transactions Adjustment
                if r >= 88 && r <= 94 && cc == 3
                  cltv_key = value.split.last 
                  @adjustment_hash[primary_key][secondry_key][cltv_key] = {}
                end
                if r >= 88 && r <= 94 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    cltv_data = 0
                  elsif @cltv_data[cc-2].include?("-")
                    cltv_data = @cltv_data[cc-2].split("-").last
                  elsif @cltv_data[cc-2].include?("≥")
                    cltv_data = @cltv_data[cc-2].split("≥").last
                  end
                  @adjustment_hash[primary_key][secondry_key][cltv_key][cltv_data] = {}
                  @adjustment_hash[primary_key][secondry_key][cltv_key][cltv_data] = value
                end

                # State Adjustments
                if r == 99 && cc == 3
                  cltv_key = value
                  @adjustment_hash[primary_key][secondry_key][cltv_key] = {}
                end
                if r ==99 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    cltv_data = 0
                  elsif @cltv_data[cc-2].include?("-")
                    cltv_data = @cltv_data[cc-2].split("-").last
                  elsif @cltv_data[cc-2].include?("≥")
                    cltv_data = @cltv_data[cc-2].split("≥").last
                  end
                  @adjustment_hash[primary_key][secondry_key][cltv_key][cltv_data] = {}
                  @adjustment_hash[primary_key][secondry_key][cltv_key][cltv_data] = value
                end
              end
            end
            (15..25).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Loan Amount Adjustments" || value == "Feature Adjustments" || value == "Product Adjustments" || value == "Max Price"
                  adj_key = value                  
                  @adjustment_hash[primary_key][adj_key] = {}
                end
                # Loan Amount Adjustments
                if r >= 67 && r <= 70 && cc == 15
                  cltv_key = value.split.last 
                  @adjustment_hash[primary_key][adj_key][cltv_key] = {}
                end
                if r >= 67 && r <= 70 && cc >15 && cc <= 25
                  if @cltv_data2[cc-2].include?("≤")
                    cltv_data = 0
                  elsif @cltv_data2[cc-2].include?("-")
                    cltv_data = @cltv_data2[cc-2].split("-").last
                  elsif @cltv_data2[cc-2].include?("≥")
                    cltv_data = @cltv_data2[cc-2].split("≥").last
                  end
                  @adjustment_hash[primary_key][adj_key][cltv_key][cltv_data] = {}
                  @adjustment_hash[primary_key][adj_key][cltv_key][cltv_data] = value
                end

                # Feature Adjustments
                if r >= 75 && r <= 80 && cc == 15
                  cltv_key = value 
                  @adjustment_hash[primary_key][adj_key][cltv_key] = {}
                end
                if r >= 75 && r <= 80 && cc >15 && cc <= 25
                  if @cltv_data2[cc-2].include?("≤")
                    cltv_data = 0
                  elsif @cltv_data2[cc-2].include?("-")
                    cltv_data = @cltv_data2[cc-2].split("-").last
                  elsif @cltv_data2[cc-2].include?("≥")
                    cltv_data = @cltv_data2[cc-2].split("≥").last
                  end
                  @adjustment_hash[primary_key][adj_key][cltv_key][cltv_data] = {}
                  @adjustment_hash[primary_key][adj_key][cltv_key][cltv_data] = value
                end

               # Product Adjustments
                if r >= 85 && r <= 90 && cc == 15
                  cltv_key = value 
                  @adjustment_hash[primary_key][adj_key][cltv_key] = {}
                end
                if r >= 85 && r <= 90 && cc >15 && cc <= 25
                  if @cltv_data2[cc-2].include?("≤")
                    cltv_data = 0
                  elsif @cltv_data2[cc-2].include?("-")
                    cltv_data = @cltv_data2[cc-2].split("-").last
                  elsif @cltv_data2[cc-2].include?("≥")
                    cltv_data = @cltv_data2[cc-2].split("≥").last
                  end
                  @adjustment_hash[primary_key][adj_key][cltv_key][cltv_data] = {}
                  @adjustment_hash[primary_key][adj_key][cltv_key][cltv_data] = value
                end

                # Max Price Adjustment
                if r >= 96 && r <= 98 && cc == 16
                  if value.include?("≤")
                    cltv_key = 0
                  elsif value.include?(">")
                    cltv_key = value.split(">").last
                  end
                  @adjustment_hash[primary_key][adj_key][cltv_key] = {}
                end
                if r >= 96 && r <= 98 && cc >16 && cc <= 23
                  if @max_price_data[cc-2].include?("30")
                    cltv_data = 30
                  elsif @max_price_data[cc-2].include?("15")
                    cltv_data = 15
                  elsif @max_price_data[cc-2].include?("10/1")
                    cltv_data = "10/1"
                  elsif @max_price_data[cc-2].include?("7/1")
                    cltv_data = "7/1"
                  elsif @max_price_data[cc-2].include?("5/1")
                    cltv_data = "5/1"
                  end
                  @adjustment_hash[primary_key][adj_key][cltv_key][cltv_data] = {}
                  @adjustment_hash[primary_key][adj_key][cltv_key][cltv_data] = value
                end
              end
            end
          end
        end
        Adjustment.create(data: @adjustment_hash,program_title: @program.title, sheet_name: sheet, program_ids: @programs_ids)
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

        @adjustment_hash = {}
        term_key = ''
        rate_type_key = ''
        jumbo_key = ''
        primary_key = ''
        (38..62).each do |r|
          row = sheet_data.row(r)
          if row.compact.count >= 1
            (0..14).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if (r == 38 && value.include?("20/25/30")) && (value.include?("Fixed") && value.include?("Jumbo")) || (r == 53 && value.include?("15") && value.include?("Fixed")) && (value.include?("Jumbo") && value.include?("ARM"))
                  if value.include?("20/25/30")
                    term_key = "20/25/30"
                  elsif value.include?("15")
                    term_key = "15"
                  end
                  rate_type_key = "Fixed"
                  jumbo_key = "Jumbo"
                  @adjustment_hash[term_key] = {}
                  @adjustment_hash[term_key][rate_type_key] = {}
                  @adjustment_hash[term_key][rate_type_key][jumbo_key] = {}
                end
                if r == 39 && cc >= 4
                  @adjustment_hash[term_key][rate_type_key][jumbo_key][dream_big_adjustment[cc].values.first] = {}
                end
                if (r > 39 && r <= 51 && r != 49 && cc >= 4)
                  @adjustment_hash[term_key][rate_type_key][jumbo_key][dream_big_adjustment[cc].values.first][dream_big_adjustment[:rows][r].values.first] = value
                end
                if r == 54  && cc >= 4
                  @adjustment_hash[term_key][rate_type_key][jumbo_key][dream_big_adjustment[:arm_column][cc].values.first] = {}
                end
                if r > 54 && r <= 62 && cc >= 4
                  @adjustment_hash[term_key][rate_type_key][jumbo_key][dream_big_adjustment[:arm_column][cc].values.first][dream_big_adjustment[:rows][r].values.first] = value
                end
              end
            end
            (0..18).each do |max_column|
              cc = 16 + max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Rate Adjustments (Increase to rate)" || value == "Max Price"
                  @key = value
                  @adjustment_hash[@key] = {}
                end
                if value == "20/25/30 Yr Fixed Only"
                  term_key = "20/25/30"
                  @adjustment_hash[@key][term_key] = {}
                end
                if value == "Rate Fall-Out Pricing Special" || value == "ARM Info"
                  @key = value
                  @adjustment_hash[@key] = {}
                end
                if r > 39 && r <= 41 && cc == 16
                  primary_key = value
                  @adjustment_hash[@key][term_key][primary_key] = {}
                elsif r > 42 && r <= 48 && cc == 16
                  primary_key = value
                  @adjustment_hash[@key][primary_key] = {}
                end
                if r > 39 && r <= 41 && cc == 18
                  @adjustment_hash[@key][term_key][primary_key] = value
                elsif r > 42 && r<= 48 && cc == 18
                  @adjustment_hash[@key][primary_key] = value
                end
                if r == 54 && cc == 16
                  primary_key = value
                  @adjustment_hash[@key][value] = {}
                end
                if r == 54 && cc == 18
                  @adjustment_hash[@key][primary_key] = value
                end
                if r > 56 && r <= 58  && cc == 16
                  primary_key = value
                  @adjustment_hash[@key][primary_key] = {}
                end
                if r > 56 && r <= 58 && cc == 17
                   @adjustment_hash[@key][primary_key] = value
                end
              end
            end
          end
        end
        make_adjust(@adjustment_hash, @program.title, sheet, @program.id)
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
        @programs_ids = []
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
              @programs_ids << @program.id
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
        Adjustment.create(data: @adjustment_hash,program_title: @program.title, sheet_name: sheet, program_ids: @programs_ids)
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def freddie_arms
    program_ids = []
    @allAdjustments = {}
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
              program_ids << @program.id
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, interest_subtype: @interest_subtype)
              @program.adjustments.destroy_all
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

        # adjustments
        previous_title = nil
        @another_title = nil
        modified_keys  = get_table_keys
        data = get_table_keys
        (49..96).each do |r|
          row    = sheet_data.row(r)
          # r == 52 / 68 / 81 / 84 / 89 / 94
          rr = r #+ 1 # (r == 53) / (r == 69) / (r == 82) / (r == 90) / (r == 95)
          max_column_section = row.compact.count - 1
          (0..max_column_section).each do |max_column|
            cc = 3 + max_column * 9 # (2 / 11)
            @title = sheet_data.cell(r,cc)
            @block_hash = {}
            if(@title.eql?("All Conforming ARMs (Does not include LP Open Access)"))
              @block_hash[@title] = {}
              key = ''
              another_key = ''
              keyOfHash = ''
              # for Misc Adjusters
              first_key  = ''
              second_key = ''
              third_key  = ''
              (0..47).each do |max_row|
                @data = []
                (7..19).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = index
                  value = sheet_data.cell(rrr,ccc)
                  # implementation of first key
                  if rrr.eql?(63)
                    # for 1st and 2nd table
                    @title = sheet_data.cell(rrr,cc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(69)
                    # for 3rd and 4th table
                    previous_title = @title = sheet_data.cell(rrr,ccc - 4) unless previous_title == @title
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["LPMI/PremiumType/FICO"] = {}
                      @block_hash[@title]["LPMI/Term/LTV/FICO"]    = {}
                    end
                  elsif rrr.eql?(81) && index == 7
                    # for Number Of Units
                    @title = sheet_data.cell(rrr,ccc - 4)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                    @another_title = "Loan Size Adjustments"
                    @block_hash[@another_title] = {} unless @block_hash.has_key?(@another_title)
                  elsif rrr.eql?(84) && index.eql?(7)
                    # for Subordinate Financing
                    @title = sheet_data.cell(rrr,ccc - 4)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(89) && index.eql?(7)
                    # for Misc Adjusters
                    @title = sheet_data.cell(rrr,ccc - 4)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(91) && index.eql?(13)
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  end

                  # implementation of second key inside first key
                  if rrr < 69 && index == 7 && value
                    # for 1st and 2nd table
                    key = get_value(value)
                    @block_hash[@title][key] = {} unless @block_hash[@title].has_key?(key)
                  elsif (69..79).to_a.include?(rrr) && index == 7 && value
                    if(rrr <= 79)
                      # for 3rd and 4th table (69..74).to_a (76..79).to_a
                      key = sheet_data.cell(rrr,ccc - 2)
                      key = get_value(key)
                      if key
                        @block_hash[@title]["LPMI/PremiumType/FICO"][key] = {} if (69..74).to_a.include?(rrr)
                        @block_hash[@title]["LPMI/Term/LTV/FICO"][key] = {} if (76..79).to_a.include?(rrr)
                      end
                    end
                  else
                    if [82,83].include?(rrr) && ccc == 7
                      # for Number Of Units
                      key = sheet_data.cell(rrr,ccc - 1)
                      @block_hash[@title][key] = {}
                    elsif rrr.eql?(81) && [18,19].include?(ccc)
                      # for Loan Size Adjustments
                      another_key = sheet_data.cell(rrr,ccc)
                      another_key = get_value(another_key)
                      @block_hash[@another_title][another_key] = {} unless @block_hash[@another_title].has_key?(another_key)
                    end

                    if (85..88).to_a.include?(rrr) && ccc < 13
                      # for Subordinate Financing
                      if index.eql?(7)
                        key = sheet_data.cell(rrr,ccc - 1)
                        key = get_value(key)
                        @block_hash[@title][key] = {} unless @block_hash[@title].has_key?(key)
                      elsif index.eql?(8)
                        keyOfHash = sheet_data.cell(rrr,ccc - 1)
                        keyOfHash = get_value(keyOfHash)
                        @block_hash[@title][key][keyOfHash] = {}
                      end
                    end

                    if (89..93).to_a.include?(rrr) && ccc < 13
                      # for Misc Adjusters
                      if index.eql?(7)
                        key = sheet_data.cell(rrr,ccc - 1)
                        if key && key.eql?("Condo > 75 LTV (>15yr Term)")
                          first_key = key.split(" >")[0]
                          @block_hash[@title][first_key] = {}
                          second_key = sheet_data.cell(rrr,ccc - 1).split(" ")[2] + ".01"
                          @block_hash[@title][first_key][second_key] = {}
                          third_key = sheet_data.cell(rrr,ccc - 1).split(" ")[4].split("(>")[1].split("yr")[0] + ".01"
                        elsif key && key.eql?(">90 LTV")
                          first_key  = key.split(" ")[1]
                          @block_hash[@title][first_key] = {}
                          second_key = key.split(">")[1].split(" ").first
                        else
                          @block_hash[@title][key] = {}
                        end
                      end
                    end

                    if (91..94).to_a.include?(rrr)
                      # for Super Conforming
                      if index.eql?(16)
                        key = sheet_data.cell(rrr,ccc)
                        @block_hash[@title][key] = {}
                      end
                    end
                  end

                  # implementation of third key inside second key with value
                  if rrr < 69 && index > 7 && value
                    # for 1st and 2nd table
                    hash_key = sheet_data.cell(rrr - (max_row + 1),ccc)
                    hash_key = get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title][key][hash_key] = value unless @block_hash[@title][key].has_key?(hash_key)
                    end
                  elsif rrr >= 69 && index >= 7 && value
                    if(rrr <= 79)
                      # for 3rd and 4th table (69..74).to_a (76..79).to_a
                      diff_of_row = rrr - 68
                      hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                      hash_key = get_value(hash_key)
                      if hash_key.present?
                        @block_hash[@title]["LPMI/PremiumType/FICO"][key][hash_key] = value if (69..74).to_a.include?(rrr)
                        @block_hash[@title]["LPMI/Term/LTV/FICO"][key][hash_key] = value if (76..79).to_a.include?(rrr)
                      end
                    elsif [82,83].include?(rrr) && [9,10,11].include?(ccc)
                      # for Number Of Units
                      diff_of_row = rrr - 81
                      hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                      hash_key = get_value(hash_key)
                      @block_hash[@title][key][hash_key] = value if hash_key.present?
                    elsif (82..89).to_a.include?(rrr) && ccc > 15 && value
                      # for Loan Size Adjustments
                      if ccc.eql?(18)
                        diff_of_column = ccc - 15
                        extra_key = sheet_data.cell(rrr,(ccc-diff_of_column))
                        @block_hash[@another_title]["Purchase"][extra_key] = value
                      else
                        diff_of_column = ccc - 15
                        extra_key = sheet_data.cell(rrr,(ccc-diff_of_column))
                        @block_hash[@another_title]["Refinance"][extra_key] = value
                      end
                    end

                    if (85..88).to_a.include?(rrr) && [10,11].include?(ccc)
                      # for Subordinate Financing
                      diff_of_row = rrr - 84
                      hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                      hash_key = get_value(hash_key)
                      @block_hash[@title][key][keyOfHash][hash_key] = value if hash_key.present?
                    end

                    if (89..93).to_a.include?(rrr) && ccc.eql?(11)
                      # for Misc Adjusters
                      if rrr.eql?(89)
                        @block_hash[@title][first_key][second_key][third_key] = value
                      elsif rrr.eql?(90)
                        @block_hash[@title][first_key][second_key] = value
                      else
                        @block_hash[@title][key] = value
                      end
                    end

                    if (91..94).to_a.include?(rrr)
                      # for Super Conforming
                      if index.eql?(19)
                        hash_key = get_value(key)
                        @block_hash[@title][key] = value if key.present?
                      end
                    end
                  end
                end

                @allAdjustments[@title] = @block_hash[@title]
                if @another_title
                  @allAdjustments[@another_title] = @block_hash[@another_title]
                end
              end
            end
          end
        end
      end
    end

    # rename first level keys
    @allAdjustments.keys.each do |key|
      data = get_table_keys
      if data[key]
        @allAdjustments[data[key]] = @allAdjustments.delete(key)
      end
    end

    # create adjustment for each program
    make_adjust(@allAdjustments, program_ids)

    redirect_to programs_import_file_path(@bank)
  end

  def conforming_arms
    program_ids = []
    @allAdjustments = {}
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
              program_ids << @program.id
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, jumbo_high_balance: @jumbo_high_balance)
              @program.adjustments.destroy_all
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

        # adjustments
        previous_title = nil
        @another_title = nil
        modified_keys  = get_table_keys
        data = get_table_keys
        (49..95).each do |r|
          row    = sheet_data.row(r)
          # r == 52 / 68 / 81 / 84 / 89 / 94
          rr = r #+ 1 # (r == 53) / (r == 69) / (r == 82) / (r == 90) / (r == 95)
          max_column_section = row.compact.count - 1
          (0..max_column_section).each do |max_column|
            cc = 3 + max_column * 9 # (2 / 11)
            @title = sheet_data.cell(r,cc)
            @block_hash = {}
            if(@title.eql?("All Conforming ARMs (Does not include DU Refi Plus)"))
              @block_hash[@title] = {}
              key = ''
              another_key = ''
              keyOfHash = ''
              # for Misc Adjusters
              first_key  = ''
              second_key = ''
              third_key  = ''
              (0..47).each do |max_row|
                @data = []
                (7..19).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = index
                  value = sheet_data.cell(rrr,ccc)
                  # implementation of first key
                  if rrr.eql?(63)
                    # for 1st and 2nd table
                    @title = sheet_data.cell(rrr,cc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(68)
                    # for 3rd table
                    previous_title = @title = sheet_data.cell(rrr,ccc - 4) unless previous_title == @title
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["LPMI/PremiumType/FICO"] = {}
                      @block_hash[@title]["LPMI/Term/LTV/FICO"]    = {}
                    end
                  elsif rrr.eql?(80) && index == 7
                    # for Loan Size Adjustments
                    @another_title = "Loan Size Adjustments"
                    @block_hash[@another_title] = {} unless @block_hash.has_key?(@another_title)
                  elsif rrr.eql?(81) && index == 7
                    # for Subordinate Financing
                    @title = sheet_data.cell(rrr,ccc - 4)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(87) && index.eql?(7)
                    # for Non Owner Occupied
                    @title = sheet_data.cell(rrr,ccc - 4)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(89) && index.eql?(13)
                    #for High Balance
                    @another_title = sheet_data.cell(rrr,ccc)
                    @block_hash[@another_title] = {} unless @block_hash.has_key?(@another_title)
                  elsif rrr.eql?(91) && index.eql?(7)
                    # for Misc Adjusters
                    @title = sheet_data.cell(rrr,ccc - 4)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  end

                  # implementation of second key inside first key
                  if rrr < 67 && index == 7 && value
                    # for 1st and 2nd table
                    key = get_value(value)
                    @block_hash[@title][key] = {} unless @block_hash[@title].has_key?(key)
                  elsif (68..79).to_a.include?(rrr) && index == 7 && value
                    if(rrr <= 79)
                      # for 3rd and 4th table (69..74).to_a (76..79).to_a
                      key = sheet_data.cell(rrr,ccc - 2)
                      key = get_value(key)
                      if key
                        @block_hash[@title]["LPMI/PremiumType/FICO"][key] = {} if (68..72).to_a.include?(rrr)
                        @block_hash[@title]["LPMI/Term/LTV/FICO"][key] = {} if (75..78).to_a.include?(rrr)
                      end
                    end
                  else
                    if (81..84).to_a.include?(rrr) && ccc < 12
                      # for Subordinate Financing
                      if index.eql?(7)
                        key = sheet_data.cell(rrr,ccc - 2)
                        key = get_value(key)
                        @block_hash[@title][key] = {} unless @block_hash[@title].has_key?(key)
                      elsif index.eql?(8)
                        keyOfHash = sheet_data.cell(rrr,ccc - 2)
                        keyOfHash = get_value(keyOfHash)
                        @block_hash[@title][key][keyOfHash] = {}
                      end
                    end

                    if rrr.eql?(80) && [18,19].include?(ccc)
                      # for Loan Size Adjustments
                      another_key = sheet_data.cell(rrr,ccc)
                      another_key = get_value(another_key)
                      @block_hash[@another_title][another_key] = {} unless @block_hash[@another_title].has_key?(another_key)
                    end

                    if [87,88,89].include?(rrr) && [7].include?(ccc)
                      #for Non Owner Occupied
                      diff_of_column = ccc - 6
                      hash_key = sheet_data.cell(rrr,(ccc -diff_of_column))
                      hash_key = get_value(hash_key)
                      key = hash_key
                      @block_hash[@title][hash_key] = {} if hash_key.present?
                    end

                    if [89,91].include?(rrr) && @another_title
                      # for High Balance
                      if index.eql?(16)
                        another_key = sheet_data.cell(rrr,ccc)
                        @block_hash[@another_title][another_key] = {} if another_key
                      end
                    end

                    if (91..95).to_a.include?(rrr) && ccc < 10
                      # for Misc Adjusters
                      if index.eql?(7)
                        key = sheet_data.cell(rrr,ccc - 2)
                        @block_hash[@title][key] = {}
                      end
                    end
                  end

                  # implementation of third key inside second key with value
                  if rrr < 67 && index > 7 && value
                    # for 1st and 2nd table
                    hash_key = sheet_data.cell(rrr - (max_row + 1),ccc)
                    hash_key = get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title][key][hash_key] = value unless @block_hash[@title][key].has_key?(hash_key)
                    end
                  elsif rrr >= 68 && index >= 7 && value
                    if(rrr <= 78)
                      # for 3rd and 4th table (69..74).to_a (76..79).to_a
                      diff_of_row = rrr - 67
                      hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                      hash_key = get_value(hash_key)
                      if hash_key.present?
                        @block_hash[@title]["LPMI/PremiumType/FICO"][key][hash_key] = value if (68..72).to_a.include?(rrr)
                        @block_hash[@title]["LPMI/Term/LTV/FICO"][key][hash_key] = value if (75..78).to_a.include?(rrr)
                      end
                    elsif (81..88).to_a.include?(rrr) && ccc > 15 && value
                      #for Loan Size Adjustments
                      if ccc.eql?(18)
                        diff_of_column = ccc - 15
                        extra_key = sheet_data.cell(rrr,(ccc-diff_of_column))
                        @block_hash[@another_title]["Purchase"][extra_key] = value
                      else
                        diff_of_column = ccc - 15
                        extra_key = sheet_data.cell(rrr,(ccc-diff_of_column))
                        @block_hash[@another_title]["Refinance"][extra_key] = value
                      end
                    end

                    if (81..84).to_a.include?(rrr) && [9,10].include?(ccc)
                      # for Subordinate Financing
                      diff_of_row = rrr - 80
                      hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                      hash_key = get_value(hash_key)
                      @block_hash[@title][key][keyOfHash][hash_key] = value if hash_key.present?
                    end

                    if [87,88,89].include?(rrr) && [9].include?(ccc)
                      @block_hash[@title][key] = value if key && value
                    end

                    if (89..92).to_a.include?(rrr)
                      # for High Balance
                      if index.eql?(19)
                        has_key = sheet_data.cell(rrr,ccc - 1)
                        @block_hash[@another_title][another_key][has_key] = value if another_key.present?
                      end
                    end

                    if (91..95).to_a.include?(rrr) && ccc.eql?(9)
                      # for Misc Adjusters
                      if rrr.eql?(89)
                        @block_hash[@title][first_key][second_key][third_key] = value
                      elsif rrr.eql?(90)
                        @block_hash[@title][first_key][second_key] = value
                      else
                        @block_hash[@title][key] = value
                      end
                    end
                  end
                end

                @allAdjustments[@title] = @block_hash[@title]
                if @another_title
                  @allAdjustments[@another_title] = @block_hash[@another_title]
                end
              end
            end
          end
        end
      end
    end

    # rename first level keys
    @allAdjustments.keys.each do |key|
      data = get_table_keys
      if data[key]
        @allAdjustments[data[key]] = @allAdjustments.delete(key)
      end
    end

    # create adjustment for each program
    make_adjust(@allAdjustments, program_ids)

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
    @programs = @bank.programs.where(sheet_name: params[:sheet])
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
      5 => {"<=80" => "0"},
      6 => {"80.01 - 85" => "80.01"},
      7 => {"> 85" => "85"},
      9 => {"< 720" => "0"},
      10 => {">= 720" => "720"},
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
        45 => {"> 85 "=> "85"},
        48 => {"<=75" => "0"},
        49 => {"<=65" => "0"},
        50 => {"65.01-75" => "65.01"},
        51 => {"75.01-80" => "75.01"},
        52 => {"80.01-90" => "80.01"},
        53 => {"90.01-95" => "90.01"},
        54 => {"All" => "All"},
        57 => {"2 Units" => "2 Unit"},
        58 => {"3-4 units" => "3-4 Unit"},
        61 => {"<$50,000" => "0"},
        62 => {"$50,000 - $99,999" => "$50,000"},
        63 => {"$100,000 - $149,999" => "$100,000"},
        64 => {"$150,000 - $199,999" => "$150,000"},
        65 => {"$200,000 - $249,999" => "$200,000"},
        66 => {"$250,000 - $299,999" => "$250,000"},
        67 => {"$300,000 - Conforming Limit" => "$300,000"},
      },
      cltv: {
        48 => {"<=80" => "0"},
        49 => {"80.01 - 95" => "80.01"},
        50 => {"80.01 - 95" => "80.01"},
        51 => {"76.01 - 95" => "76.01"},
        52 => {"81.01 - 95" => "81.01"},
        53 => {"91.01 - 95" => "91.01"},
        54 => {"> 95" => "95"}
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
          28 => {">=760" => "760"},
          29 => {"740-759" => "740"},
          30 => {"720-739" => "720"},
          31 => {"700-719" => "700"},
          32 => {"680-699" => "680"},
          34 => {">=760" => "760"},
          35 => {"740-759" => "740"},
          36 => {"720-739" => "720"},
          37 => {"700-719" => "700"},
          38 => {"680-699" => "680"},
        },
        subordinate: {
          4 => {"< 720" => "0"},
          5 => {">= 720" => "720"}
        }
      }
    return data
  end

  def jumbo_series_i_adjustment
      data = {
        5 => {"≤ 60" => "0"},
        6 => {"60.01-65" => "60.01"},
        7 => {"65.01-70" => "65.01"},
        8 => {"70.01-75" => "70.01"},
        10 => {"75.01-80" => "75.01"},
        14 => {"≤ 60" => "0"},
        16 => {"60.01-65" => "60.01"},
        17 => {"65.01-70" => "65.01"},
        18 => {"70.01-75" => "70.01"},
        19 => {"75.01-80" => "75.01"},
        rows: {
          41 => {"< 700" => "0"},
          42 => {"740-759" => "740"},
          43 => {"720-739" => "720"},
          44 => {"700-719" => "700"},
          45 => {"680-699" => "680"},
          50 => {"≤ $1MM" => "0"},
          51 => {"$1MM - $1.5MM" => "$1MM"},
          52 => {"$1.5MM - $2MM" => "$1.5MM"},
          53 => {"$2MM - $2.5MM" => "$2MM"},
          58 => {"2nd Home" => "2nd Home"},
          59 => {"Purchase (15 Yr Fixed ONLY)" => "Purchase (15 Yr Fixed ONLY)"},
          60 => {"C/O Refinance" => "C/O Refinance"},
          61 => {"2-4 Unit" => "2-4 Unit"},
          62 => {"DTI > 40%" => "DTI > 40%"}
        }
      }
    return data
  end

  def dream_big_adjustment
    data = {
      4 => {"<=50" => "0"},
      5 => {"50.01 - 55" => "50.01"},
      6 => {"55.01 - 60" => "55.01"},
      7 => {"60.01 - 65" => "60.01"},
      9 => {"65.01 - 70" => "65.01"},
      10 => {"70.01 - 75" => "70.01"},
      11 => {"75.01 - 80" => "75.01"},
      12 => {"80.01 - 85" => "80.01"},
      14 => {"85.01 - 90" => "85.01"},
      rows: {
        40 => {"680 - 699" => "680"},
        41 => {"700 - 719" => "700"},
        42 => {"720 - 739" => "720"},
        43 => {"740 - 759" => "740"},
        44 => {"760-779" => "760"},
        45 => {">=780"=> "780"},
        46 => {"Purchase" => "Purchase"},
        47 => {"Cash Out Refinance" => "Cash Out Refinance"},
        48 => {"Rate & Term Refinance" => "Rate & Term Refinance"},
        50 => {"Non Owner Occupied" => "Non Owner Occupied"},
        51 => {"> 80 LTV No MI" => "> 80 LTV No MI"},
        55 => {"680 - 699" => "680"},
        56 => {"700 - 719" => "700"},
        57 => {"720 - 739" => "720"},
        58 => {"740 - 759" => "740"},
        59 => {"760-779" => "760"},
        60 => {">=780"=> "780"},
        61 => {"Purchase" => "Purchase"},
        62 => {"Cash Out Refinance" => "Cash Out Refinance"},
      },
      arm_column: {
        4 => {"<=50" => "0"},
        5 => {"50.01 - 55" => "50.01"},
        6 => {"55.01 - 60" => "55.01"},
        8 => {"60.01 - 65" => "60.01"},
        9 => {"65.01 - 70" => "65.01"},
        10 => {"70.01 - 75" => "70.01"},
        11 => {"75.01 - 80" => "75.01"},
        12 => {"80.01 - 85" => "80.01"},
        14 => {"85.01 - 90" => "85.01"},
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

  def make_adjust(block_hash, p_ids)
    begin
      adjustment = Adjustment.create(data: block_hash.to_json)

      # assign for all projects
      p_ids.each do |id|
        program = Program.find(id)
        program.adjustments << adjustment
      end
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



  def get_table_keys
    table_keys = {
      "All Conforming ARMs (Does not include LP Open Access)" => "Conforming/Term/LTV/FICO",
      "Cash-Out" => "RefinanceType/LTV/FICO",
      "Lender Paid MI Adjustments" => "Term/LTV/FICO",
      "Premium Adjustments" => "LPMI/PremiumType/FICO",
      "LTV Adjustments" => "LPMI/Term/LTV/FICO",
      "Number Of Units" => "PropertyType/LTV",
      "Subordinate Financing" => "FinancingType/LTV/CLTV/FICO",
      "Misc Adjusters" => "PropertyType/LTV/Term",
      "Non Owner Occupied" => "PropertyType/LTV",
      "Loan Size Adjustments" => "RateType/Conforming/LTV/FICO",
      "Super Conforming" => "Conforming/LTV/FICO",
    }

    return table_keys
  end


  def get_value value1
    if value1.present?
      if (value1.include?("≤")) || (value1.include?("<"))
        value1 = 0
      elsif (value1.include?("-"))
        value1 = value1.split("-").first.squish
      elsif (value1.include?("≥"))
        value1 = value1.split("≥").last.squish
      elsif (value1.include?(">="))
        value1.split(">=").last.squish
      elsif (value1.include?(">"))
        value1.split(">").last.squish
      elsif (value1.include?("+"))
        value1.split("+").first
      else
        value1
      end
    end
  end
end
