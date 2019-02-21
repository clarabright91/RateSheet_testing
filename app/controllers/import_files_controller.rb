class ImportFilesController < ApplicationController
  # before_action :get_bank, only: [:import_government_sheet, :programs, :import_freddie_fixed_rate, :import_conforming_fixed_rate, :home_possible, :conforming_arms, :lp_open_acces_arms, :lp_open_access_105, :lp_open_access, :du_refi_plus_arms, :du_refi_plus_fixed_rate_105, :du_refi_plus_fixed_rate, :dream_big, :high_balance_extra, :freddie_arms, :jumbo_series_d,:jumbo_series_f, :jumbo_series_h, :jumbo_series_i, :jumbo_series_jqm, :import_homereddy_sheet, :import_HomeReadyhb_sheet]
  before_action :get_sheet, only: [:government, :programs, :freddie_fixed_rate, :conforming_fixed_rate, :home_possible, :conforming_arms, :lp_open_acces_arms, :lp_open_access_105, :lp_open_access, :du_refi_plus_arms, :du_refi_plus_fixed_rate_105, :du_refi_plus_fixed_rate, :dream_big, :high_balance_extra, :freddie_arms, :jumbo_series_d,:jumbo_series_f, :jumbo_series_h, :jumbo_series_i, :jumbo_series_jqm, :homeready, :homeready_hb]
  before_action :get_program, only: [:single_program, :program_property]
  require 'roo'
  require 'roo-xls'

  def index
    # HardWorker.perform_async(1)
    @banks = Bank.all
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @sheetlist =[]
    begin
      xlsx.sheets.each do |sheet|
        @sheetlist.push(sheet)
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
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue Exception => e
      # the required headers are not all present
    end
  end

  def government
    @programs_ids = []
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
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

              @title = sheet_data.cell(r,cc)
              # term
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              else
                term = nil
              end

               # rate arm
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # streamline && fha, Va , Usda
              fha = false
              va = false
              usda = false
              streamline = false
              full_doc = false
              if @title.include?("FHA")
                streamline = true
                fha = true
                full_doc = true
              elsif @title.include?("VA")
                streamline = true
                va = true
                full_doc = true
              elsif @title.include?("USDA")
                streamline = true
                usda = true
                full_doc = true
              else
                streamline = false
                fha = false
                va = false
                usda = false
                full_doc = false
              end

               # rate arm
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              # High Balance
              jumbo_high_balance = false
              if @title.include?("High Balance")
                jumbo_high_balance = true
              end
              
              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @programs_ids << @program.id
                # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",streamline: streamline, fha: fha, va: va, usda: usda, full_doc: full_doc, jumbo_high_balance: jumbo_high_balance,sheet_name: sheet, arm_basic: arm_basic)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              if @program.fha
                gov_key = "FHA"
              elsif @program.va
                gov_key = "VA"
              elsif @program.usda
                gov_key = "USDA"
              end
              if @program.term.present?
                term = @program.term
              end
              if @program.loan_type.present?
                loan_type = @program.loan_type
              end

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
                  main_key = "LoanType/FICO/LTV"
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
                  main_key = "LoanPurpose/LoanAmount/LTV"
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
                  main_key = "VA/LoanPurpose/LoanAmount/LTV"
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
    # create_program_association_with_adjustment(sheet)
    # redirect_to programs_import_file_path(@bank)
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def freddie_fixed_rate
    @program_ids = []
    @allAdjustments = {}
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Freddie Fixed Rate")
        sheet_data = xlsx.sheet(sheet)
        @sheet = sheet
        main_key = ''
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
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # interest type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # conforming
              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              # freddie_mac
              freddie_mac = false
              if @title.include?("Freddie Mac")
                freddie_mac = true
              end

              # fannie_mae
              fannie_mae = false
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                fannie_mae = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
                # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming,freddie_mac: freddie_mac, sheet_name: sheet, fannie_mae: fannie_mae)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    # first_row[c_i]
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
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
              @title = "LoanSize/LoanType/Term/FICO/LTV"
              @block_hash[@title] = {}
              @block_hash[@title]["Conforming"] = {}
              @block_hash[@title]["Conforming"]["Fixed"] = {}
              @block_hash[@title]["Conforming"]["Fixed"]["0-15"] = {}
              @block_hash["PropertyType"] = {} unless @block_hash.has_key?("PropertyType")
              @block_hash["MiscAdjuster"] = {} unless @block_hash.has_key?("MiscAdjuster")
              @block_hash["LockDay"] = {} unless @block_hash.has_key?("LockDay")

              key = ''
              another_key = ''
              keyOfHash   = ''
              # for Misc Adjusters
              first_key   = ''
              second_key  = ''
              third_key   = ''
              final_key   = ''
              f1_key      = ''
              f2_key      = ''
              loan_amount = ''

              (0..50).each do |max_row|
                @data = []
                (3..19).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = index
                  value = sheet_data.cell(rrr,ccc)
                  # implementation of first key
                  if rrr.eql?(132)
                    # for Cash-Out
                    @title = sheet_data.cell(rrr,cc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Cash Out"] = {}
                    end
                  elsif rrr.eql?(138) && index == 3
                    # for Lender Paid MI Adjustments
                    previous_title = @title = sheet_data.cell(rrr,ccc) unless previous_title == @title
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      first_key  = "LPMI/RefinanceOption/LTV"
                      second_key = "LPMI/PropertyType/LTV"
                      final_key  = "LPMI/Term/FICO/LTV"
                      @block_hash[@title][first_key] = {}
                      @block_hash[@title][first_key][true] = {}
                      @block_hash[@title][second_key] = {}
                      @block_hash[@title][second_key][true] = {}
                      @block_hash[@title][final_key] = {}
                      @block_hash[@title][final_key][true] = {}
                    end
                  elsif rrr.eql?(155) && index == 3
                    # for Number Of Units
                    @title = sheet_data.cell(rrr,ccc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                    end
                  elsif rrr.eql?(156) && index == 13
                    # for Loan Size Adjustments
                    @another_title = sheet_data.cell(rrr,index)
                    unless @block_hash.has_key?(@another_title)
                      @block_hash[@another_title] = {}
                    end
                  elsif rrr.eql?(158) && index == 3
                    # for Subordinate Financing
                    @title = sheet_data.cell(rrr,ccc)
                    # @title = "FinancingType/LTV/CLTV/FICO"
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Subordinate Financing"] = {}
                    end
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
                    @block_hash[@another_title]["Non Owner Occupied"] = {}
                  end

                  #implementation of second key inside first key
                  if rrr > 122 && rrr < 131 && index == 7 && value
                    key = get_value(value)
                    @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key] = {} unless @block_hash[@title]["Conforming"]["Fixed"]["0-15"].has_key?(key)
                  elsif rrr > 131 && rrr < 136 && index == 7 && value
                    # for 1st and 2nd table
                    key = get_value(value)
                    @block_hash[@title]["Cash Out"][key] = {} unless @block_hash[@title]["Cash Out"].has_key?(key)
                  elsif (rrr > 137) && (rrr < 154)
                    # for Lender Paid MI Adjustments
                    if index == 5 && value
                      if ["Rate & Term Refi", "Cash Out"].include?(value)
                        f1_key = "Rate & Term"
                        @block_hash[@title][first_key][true][f1_key] = {}
                      elsif ["Manufactured Home", "2nd Home", "3-4 Unit", "Non Owner Occupied"].include?(value)
                        f2_key = value
                        @block_hash[@title][second_key][true][f2_key] = {}
                      else
                        @block_hash[@title][final_key][true]["0 - 20"] = {} if value.eql?("≤ 20 Yr Term")
                        @block_hash[@title][final_key][true]["20 - Infinity"] = {} if value.eql?("> 20 Yr Term")
                        key = "0 - 20" if value.eql?("≤ 20 Yr Term")
                        key = "20 - Infinity" if value.eql?("> 20 Yr Term")
                      end
                    elsif index == 6 && rrr < 154 && value
                      another_key = value.eql?("≤ 85") ? set_range(value) : value
                      @block_hash[@title][final_key][true][key][another_key] = {} if another_key
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
                      @block_hash[@title]["Subordinate Financing"][key] = {} unless @block_hash[@title].has_key?(key)
                    elsif index.eql?(7)
                      keyOfHash = sheet_data.cell(rrr,ccc)
                      keyOfHash = get_value(keyOfHash)
                      @block_hash[@title]["Subordinate Financing"][key][keyOfHash] = {}
                    end
                  end

                  if (rrr >= 156 && rrr <= 163) && index.eql?(15)
                    # for Loan Size Adjustments
                    loan_amount = get_value(sheet_data.cell(rrr, ccc))
                    @block_hash[@another_title][loan_amount] = {}
                  end

                  if (163..166).to_a.include?(rrr) && ccc < 10
                    # for Misc Adjusters
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      if key && key.eql?("Condo > 75 LTV (>15yr Term)")
                        first_key = key.split(" >")[0]
                        @block_hash[@title][first_key] = {}
                        second_key = sheet_data.cell(rrr,ccc).split(" ")[2] + ".01"
                        @block_hash[@title][first_key][second_key] = {}
                        third_key = sheet_data.cell(rrr,ccc).split(" ")[4].split("(>")[1].split("yr")[0] + ".01"
                      elsif key && key.eql?(">90 LTV")
                        first_key  = key.split(" ")[1]
                        @block_hash[@title][first_key] = {}
                        second_key = key.split(">")[1].split(" ").first
                      end
                    end
                  end

                  if [167,168,169].include?(rrr) && [7].include?(ccc)
                    #for Non Owner Occupied
                    hash_key = sheet_data.cell(rrr,ccc)
                    hash_key = hash_key.eql?("> 80") ? set_range(hash_key) : get_value(hash_key)
                    key = hash_key
                    @block_hash[@another_title]["Non Owner Occupied"][hash_key] = {} if hash_key.present?
                  end

                  if [164,165].include?(rrr) && @another_title
                    # for Super Conforming Adjustments
                    if index.eql?(17)
                      another_key = sheet_data.cell(rrr,ccc)
                      @block_hash[@another_title]["Purchase"] = {} unless @block_hash[@another_title].has_key?("Purchase")
                      @block_hash[@another_title]["Refinance"] = {} unless @block_hash[@another_title].has_key?("Refinance")
                    end
                  end

                  # implementation of third key inside second key with value
                  if rrr > 122 && rrr < 131 && index > 7 && value
                    # for 1st table
                    diff_of_row = rrr - 122
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 760") ? set_range(hash_key) : get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key][hash_key] = value unless @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 131 && rrr < 136 && index > 7 && value
                    # for 2nd table
                    diff_of_row = rrr - 122
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 760") ? set_range(hash_key) : get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Cash Out"][key][hash_key] = value unless @block_hash[@title]["Cash Out"][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 137 && rrr <= 153 && index >= 7 && value
                    # for Lender Paid MI Adjustments
                    diff_of_row = rrr - 137
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = set_range(hash_key).nil? ? get_value(hash_key) : set_range(hash_key)
                    if (138..143).to_a.include?(rrr)
                      if [138, 143].include?(rrr)
                        @block_hash[@title][first_key][true][f1_key][hash_key] = value
                      elsif [139, 140, 141, 142].include?(rrr)
                        @block_hash[@title][second_key][true][f2_key][hash_key] = value
                      end
                    else
                      if rrr > 144
                        @block_hash[@title][final_key][true][key][another_key][hash_key] = value
                      end
                    end
                  end

                  if [156,157].include?(rrr) && [9,10,11].include?(ccc)
                    # for Number Of Units
                    diff_of_row = rrr - 155
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = (hash_key.eql?("≤ 80") || hash_key.eql?("> 85")) ? set_range(hash_key) : get_value(hash_key)
                    @block_hash[@title][key][hash_key] = value if hash_key.present?
                  end

                  if (159..162).to_a.include?(rrr) && ccc > 9 && ccc < 12 && value
                    # for Subordinate Financing
                    diff_of_row = rrr - 158
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 720") ? set_range(hash_key) : get_value(hash_key)
                    @block_hash[@title]["Subordinate Financing"][key][keyOfHash][hash_key] = value if hash_key.present?
                  end

                  if (156..163).to_a.include?(rrr) && ccc > 15 && value
                    #for Loan Size Adjustments
                    diff_of_row  = rrr - 155
                    loan_purpose = sheet_data.cell(rrr-diff_of_row,ccc)
                    @block_hash[@another_title][loan_amount][loan_purpose] = value
                  end

                  if (163..166).to_a.include?(rrr) && ccc == 11
                    #for Misc Adjusters
                    if rrr.eql?(163)
                      @block_hash[@title][first_key][second_key][third_key] = value
                    else
                      first_key = sheet_data.cell(rrr,ccc - 5)
                      @block_hash["PropertyType"]["Manufactured Home"] = value if ["Manufactured Home (High Bal - Not Permitted)"].include?(first_key)
                      @block_hash["MiscAdjuster"][first_key] = value if ["CA Escrow Waiver (Full or Taxes Only)", "CA Escrow Waiver (Insurance Only)"].include?(first_key)
                    end
                  end

                  if [167,168,169].include?(rrr) && [11].include?(ccc)
                    #for Non Owner Occupied
                    @block_hash[@another_title]["Non Owner Occupied"][key] = value if key && value
                  end

                  if [164,165].to_a.include?(rrr)
                    # for Super Conforming Adjustments
                    if index.eql?(19)
                      has_key = sheet_data.cell(rrr,ccc)
                      if rrr.eql?(164)
                        @block_hash[@another_title]["Purchase"]["Rate and Term"] = value
                        @block_hash[@another_title]["Refinance"]["Rate and Term"] = value
                      else
                        @block_hash[@another_title]["Refinance"]["Cash Out"] = value
                      end
                    end
                  end

                  if [172,173].to_a.include?(rrr) && ccc.eql?(17)
                    if rrr.eql?(172)
                      @block_hash["LockDay"]["30"] = value
                      @block_hash["LockDay"]["45"] = value
                      @block_hash["LockDay"]["60"] = value
                    else
                      @block_hash["LockDay"]["90"] = value
                    end
                  end
                end

                @allAdjustments[@title] = @block_hash[@title]
                if @another_title
                  @allAdjustments[@another_title] = @block_hash[@another_title]
                  @allAdjustments["PropertyType"] = @block_hash["PropertyType"]
                  @allAdjustments["MiscAdjuster"] = @block_hash["MiscAdjuster"]
                  @allAdjustments["LockDay"] = @block_hash["LockDay"]
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
    make_adjust(@allAdjustments, @sheet)
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def conforming_fixed_rate
    program_ids = []
    @allAdjustments = {}
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Conforming Fixed Rate")
        @sheet = sheet
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
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # interest type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # conforming
              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              # freddie_mac
              freddie_mac = false
              if @title.include?("Freddie Mac")
                freddie_mac = true
              end

              # fannie_mae
              fannie_mae = false
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                fannie_mae = true
              end

              # High Balance
              jumbo_high_balance = false
              if @title.include?("High Balance")
                jumbo_high_balance = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              program_ids << @program.id
                # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, sheet_name: sheet,jumbo_high_balance: jumbo_high_balance)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    # first_row[c_i]
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        # adjustments
        previous_title = nil
        @another_title = nil
        modified_keys  = get_table_keys
        data = get_table_keys
        (123..171).each do |r|
          row = sheet_data.row(r)
          # r == 52 / 68 / 81 / 84 / 89 / 94
          rr = r #+ 1 # (r == 53) / (r == 69) / (r == 82) / (r == 90) / (r == 95)
          max_column_section = 0#row.compact.count - 1
          (0..max_column_section).each do |max_column|
            cc = 3 + max_column * 9 # (2 / 11)
            @title = sheet_data.cell(r,cc)
            @block_hash = {}
            if(@title.eql?("All Fixed Conforming\n(does not apply to terms ≤ 15yrs)"))
              @title = "LoanSize/LoanType/Term/FICO/LTV"
              @block_hash[@title] = {}
              @block_hash[@title]["Conforming"] = {}
              @block_hash[@title]["Conforming"]["Fixed"] = {}
              @block_hash[@title]["Conforming"]["Fixed"]["0-15"] = {}
              @block_hash["PropertyType"] = {} unless @block_hash.has_key?("PropertyType")
              @block_hash["MiscAdjuster"] = {} unless @block_hash.has_key?("MiscAdjuster")
              @block_hash["LockDay"] = {} unless @block_hash.has_key?("LockDay")
              #  for high balance key
              @block_hash["LoanSize/LoanPurpose/RefinanceOption"] = {} unless @block_hash.has_key?("LoanSize/LoanPurpose/RefinanceOption")
              @block_hash["LoanSize/LoanPurpose/RefinanceOption"]["High Balance"] = {} unless @block_hash["LoanSize/LoanPurpose/RefinanceOption"].has_key?("High Balance")
              @block_hash["LoanSize/RefinanceOption"] = {} unless @block_hash.has_key?("LoanSize/RefinanceOption")
              @block_hash["LoanSize/RefinanceOption"]["High Balance"] = {} unless @block_hash["LoanSize/RefinanceOption"].has_key?("High Balance")

              key = ''
              another_key = ''
              keyOfHash   = ''
              # for Misc Adjusters
              first_key   = ''
              second_key  = ''
              third_key   = ''
              final_key   = ''
              f1_key      = ''
              f2_key      = ''
              loan_amount = ''
              (0..48).each do |max_row|
                @data = []
                (3..19).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = index
                  value = sheet_data.cell(rrr,ccc)
                  # implementation of first key
                  if rrr.eql?(133)
                    # for Cash-Out
                    @title = sheet_data.cell(rrr,cc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Cash Out"] = {}
                    end
                  elsif rrr.eql?(138) && index == 3
                    # for Lender Paid MI Adjustments
                    previous_title = @title = sheet_data.cell(rrr,ccc) unless previous_title == @title
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      first_key  = "LPMI/RefinanceOption/LTV"
                      second_key = "LPMI/PropertyType/LTV"
                      final_key  = "LPMI/Term/FICO/LTV"
                      @block_hash[@title][first_key] = {}
                      @block_hash[@title][first_key][true] = {}
                      @block_hash[@title][second_key] = {}
                      @block_hash[@title][second_key][true] = {}
                      @block_hash[@title][final_key] = {}
                      @block_hash[@title][final_key][true] = {}
                    end
                  elsif rrr.eql?(156) && index == 3
                    # for Subordinate Financing
                    @title = sheet_data.cell(rrr,ccc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Subordinate Financing"] = {}
                    end
                  elsif rrr.eql?(156) && index == 13
                    # for Loan Size Adjustments
                    @another_title = sheet_data.cell(rrr,index)
                    unless @block_hash.has_key?(@another_title)
                      @block_hash[@another_title] = {}
                    end
                  elsif rrr.eql?(162) && index == 3
                    # for Misc Adjusters
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(167) && index.eql?(3)
                    # for Non Owner Occupied
                    @another_title = sheet_data.cell(rrr,ccc)
                    @block_hash[@another_title] = {} unless @block_hash.has_key?(@another_title)
                    @block_hash[@another_title]["Non Owner Occupied"] = {}
                  end

                  #implementation of second key inside first key
                  if rrr > 122 && rrr < 131 && index == 7 && value
                    key = get_value(value)
                    @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key] = {} unless @block_hash[@title]["Conforming"]["Fixed"]["0-15"].has_key?(key)
                  elsif rrr > 132 && rrr < 136 && index == 7 && value
                    # for 1st and 2nd table
                    key = get_value(value)
                    @block_hash[@title]["Cash Out"][key] = {} unless @block_hash[@title]["Cash Out"].has_key?(key)
                  elsif (rrr > 137) && (rrr < 154)
                    # for Lender Paid MI Adjustments
                    if index == 5 && value
                      if ["Rate & Term Refi", "Cash Out"].include?(value)
                        f1_key = "Rate & Term Refi".eql?(value) ? "Rate & Term" : "Cash Out"
                        @block_hash[@title][first_key][true][f1_key] = {}
                      elsif ["Manufactured Home", "2nd Home", "3-4 Unit", "Non Owner Occupied"].include?(value)
                        f2_key = value
                        @block_hash[@title][second_key][true][f2_key] = {}
                      else
                        @block_hash[@title][final_key][true]["0 - 20"] = {} if value.eql?("≤ 20 Yr Term")
                        @block_hash[@title][final_key][true]["20 - Infinity"] = {} if value.eql?("> 20 Yr Term")
                        key = "0 - 20" if value.eql?("≤ 20 Yr Term")
                        key = "20 - Infinity" if value.eql?("> 20 Yr Term")
                      end
                    elsif index == 6 && rrr < 154 && value
                      another_key = value.eql?("≤ 85") ? set_range(value) : value
                      @block_hash[@title][final_key][true][key][another_key] = {} if another_key
                    end
                  end

                  if (156..161).to_a.include?(rrr) && ccc < 12
                    # for Subordinate Financing
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      key = get_value(key)
                      @block_hash[@title]["Subordinate Financing"][key] = {} unless @block_hash[@title].has_key?(key)
                    elsif index.eql?(7)
                      keyOfHash = sheet_data.cell(rrr,ccc)
                      keyOfHash = get_value(keyOfHash)
                      @block_hash[@title]["Subordinate Financing"][key][keyOfHash] = {}
                    end
                  end

                  if (rrr >= 156 && rrr <= 163) && index.eql?(15)
                    # for Loan Size Adjustments
                    loan_amount = get_value(sheet_data.cell(rrr, ccc))
                    @block_hash[@another_title][loan_amount] = {}
                  end

                  if (162..166).to_a.include?(rrr)
                    # for Misc Adjusters
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      if key && key.eql?("Attached Condo > 75 LTV (>15yr Term)")
                        first_key = key.split(" >")[0].split(" ").last
                        @block_hash[@title][first_key] = {}
                        second_key = sheet_data.cell(rrr,ccc).split(" ")[3] + ".01"
                        @block_hash[@title][first_key][second_key] = {}
                        third_key = sheet_data.cell(rrr,ccc).split(" ")[5].split("(>")[1].split("yr")[0] + ".01"
                      elsif key && key.eql?(">90 LTV")
                        first_key  = key.split(" ")[1]
                        @block_hash[@title][first_key] = {}
                        second_key = key.split(">")[1].split(" ").first
                      end
                    end
                  end

                  if [167,168,169].include?(rrr) && [7].include?(ccc)
                    #for Non Owner Occupied
                    hash_key = sheet_data.cell(rrr,ccc)
                    unless hash_key.eql?("> 80")
                      hash_key = get_value(hash_key)
                      key = hash_key
                    else
                      hash_key = set_range(hash_key)
                      key = hash_key
                    end
                    @block_hash[@another_title]["Non Owner Occupied"][hash_key] = {} if hash_key.present?
                  end

                  # implementation of third key inside second key with value
                  if rrr > 122 && rrr < 131 && index > 7 && value
                    # for 1st table
                    diff_of_row = rrr - 122
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 760") ? set_range(hash_key) : get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key][hash_key] = value unless @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 132 && rrr < 136 && index > 7 && value
                    # for 2nd table
                    diff_of_row = rrr - 122
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 760") ? set_range(hash_key) : get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Cash Out"][key][hash_key] = value unless @block_hash[@title]["Cash Out"][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 137 && rrr <= 153 && index >= 7 && value
                    # for Lender Paid MI Adjustments
                    diff_of_row = rrr - 137
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = set_range(hash_key).nil? ? get_value(hash_key) : set_range(hash_key)
                    if (138..143).to_a.include?(rrr)
                      if [138, 143].include?(rrr)
                        @block_hash[@title][first_key][true][f1_key][hash_key] = value
                      elsif [139, 140, 141, 142].include?(rrr)
                        @block_hash[@title][second_key][true][f2_key][hash_key] = value
                      end
                    else
                      if rrr > 144
                        @block_hash[@title][final_key][true][key][another_key][hash_key] = value
                      end
                    end
                  end

                  if (156..161).to_a.include?(rrr) && ccc > 9 && ccc < 12 && value
                    # for Subordinate Financing
                    diff_of_row = rrr - 155
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 720") ? set_range(hash_key) : get_value(hash_key)
                    @block_hash[@title]["Subordinate Financing"][key][keyOfHash][hash_key] = value if hash_key.present?
                  end

                  if (156..163).to_a.include?(rrr) && ccc > 15 && value
                    #for Loan Size Adjustments
                    diff_of_row  = rrr - 155
                    loan_purpose = sheet_data.cell(rrr-diff_of_row,ccc)
                    @block_hash[@another_title][loan_amount][loan_purpose] = value
                  end

                  if (162..166).to_a.include?(rrr) && ccc == 11
                    #for Misc Adjusters
                    if rrr.eql?(164)
                      @block_hash[@title][first_key][second_key][third_key] = value
                    else
                      first_key = sheet_data.cell(rrr,ccc - 5)
                      @block_hash["PropertyType"][first_key] = value if ["2-4 Units", "Manufactured Home"].include?(first_key)
                      @block_hash["MiscAdjuster"][first_key] = value if ["CA Escrow Waiver (Full or Taxes Only)", "CA Escrow Waiver (Insurance Only)"].include?(first_key)
                    end
                  end

                  if [167,168,169].include?(rrr) && [11].include?(ccc)
                    @block_hash[@another_title]["Non Owner Occupied"][key] = value if key && value
                  end

                  if [171,172].include?(rrr)
                    # for High Balance
                    if rrr.eql?(171) && index.eql?(19)
                      @block_hash["LoanSize/LoanPurpose/RefinanceOption"]["High Balance"]["Purchase"] = {}
                      @block_hash["LoanSize/LoanPurpose/RefinanceOption"]["High Balance"]["Purchase"]["Rate and Term"] = value
                    else
                      @block_hash["LoanSize/RefinanceOption"]["High Balance"]["Cash Out"] = value
                    end
                  end

                  if [169,170].to_a.include?(rrr) && ccc.eql?(17)
                    if rrr.eql?(169)
                      @block_hash["LockDay"]["30"] = value
                      @block_hash["LockDay"]["45"] = value
                      @block_hash["LockDay"]["60"] = value
                    else
                      @block_hash["LockDay"]["90"] = value
                    end
                  end
                end

                @allAdjustments[@title] = @block_hash[@title]
                if @another_title
                  @allAdjustments[@another_title] = @block_hash[@another_title]
                  @allAdjustments["PropertyType"] = @block_hash["PropertyType"]
                  @allAdjustments["MiscAdjuster"] = @block_hash["MiscAdjuster"]
                  @allAdjustments["LockDay"] = @block_hash["LockDay"]
                  @allAdjustments["LoanSize/LoanPurpose/RefinanceOption"] = @block_hash["LoanSize/LoanPurpose/RefinanceOption"]
                  @allAdjustments["LoanSize/RefinanceOption"] = @block_hash["LoanSize/RefinanceOption"]
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
    make_adjust(@allAdjustments, @sheet)
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def home_possible
    @program_ids = []
    @allAdjustments = {}
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
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
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # rate type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              # conforming
              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              # freddie_mac
              freddie_mac = false
              if @title.include?("Freddie Mac")
                freddie_mac = true
              end

              # fannie_mae
              fannie_mae = false
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                fannie_mae = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
              # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.adjustments.destroy_all
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, sheet_name: sheet,arm_basic: arm_basic)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
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
            if(@title.eql?("All Conforming (does not apply to Fixed terms ≤ 15yrs)"))
              @block_hash["PropertyType"] = {} unless @block_hash.has_key?("PropertyType")
              @block_hash["LockDay"] = {} unless @block_hash.has_key?("LockDay")
              key = ''
              another_key = ''
              keyOfHash = ''
              # for Misc Adjusters
              first_key   = ''
              second_key  = ''
              third_key   = ''
              final_key   = ''
              f1_key      = ''
              f2_key      = ''
              loan_amount = ''
              (0..55).each do |max_row|
                @data = []
                (3..19).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = index
                  value = sheet_data.cell(rrr,ccc)
                  # implementation of first key
                  if rrr.eql?(81) && index == 3
                    # for All Conforming
                    @title = "LoanSize/LoanType/Term/FICO/LTV"
                    @block_hash[@title] = {}
                    @block_hash[@title]["Conforming"] = {}
                    @block_hash[@title]["Conforming"]["Fixed"] = {}
                    @block_hash[@title]["Conforming"]["Fixed"]["0-15"] = {}
                  elsif rrr.eql?(93) && index == 3
                    # for Lender Paid MI Adjustments
                    previous_title = @title = sheet_data.cell(rrr,ccc) unless previous_title == @title
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      first_key  = "LPMI/RefinanceOption/LTV"
                      second_key = "LPMI/PropertyType/LTV"
                      final_key  = "LPMI/Term/FICO/LTV"
                      @block_hash[@title][first_key] = {}
                      @block_hash[@title][first_key][true] = {}
                      @block_hash[@title][second_key] = {}
                      @block_hash[@title][second_key][true] = {}
                      @block_hash[@title][final_key] = {}
                      @block_hash[@title][final_key][true] = {}
                    end
                  elsif rrr.eql?(107) && index == 3
                    # for VLIP LPMI Adjustments
                    @block_hash["LPMI/PropertyType/FICO"] = {} unless @block_hash.has_key?("LPMI/PropertyType/FICO")
                    @block_hash["LPMI/Term/FICO/LTV"] = {} unless @block_hash.has_key?("LPMI/Term/FICO/LTV")
                  elsif rrr.eql?(115) && index == 3
                    # for Subordinate Financing
                    @title = sheet_data.cell(rrr,ccc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Subordinate Financing"] = {}
                    end
                  elsif rrr.eql?(115) && index == 13
                    # for Loan Size Adjustments
                    @another_title = sheet_data.cell(rrr,index)
                    unless @block_hash.has_key?(@another_title)
                      @block_hash[@another_title] = {}
                    end
                  elsif rrr.eql?(120) && index == 3
                    # for Misc Adjusters
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(123) && index.eql?(3)
                    # for Number Of Units
                    @title = sheet_data.cell(rrr,ccc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                    end
                  elsif rrr.eql?(130) && index.eql?(13)
                    # for Adjustment Caps
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  end

                  #implementation of second key inside first key
                  if (81..88).to_a.include?(rrr) && index == 7 && value
                    # for All Conforming
                    key = get_value(value)
                    if key
                      @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key] = {} unless @block_hash[@title]["Conforming"]["Fixed"]["0-15"].has_key?(key)
                    end
                  end

                  if (rrr > 92) && (rrr < 106)
                    # for Lender Paid MI Adjustments
                    if index == 5 && value
                      if ["Rate & Term Refi", "Cash Out"].include?(value)
                        f1_key = "Rate & Term Refi".eql?(value) ? "Rate & Term" : "Cash Out"
                        @block_hash[@title][first_key][true][f1_key] = {}
                      elsif ["Manufactured Home", "2nd Home", "3-4 Unit", "Non Owner Occupied"].include?(value)
                        f2_key = value
                        @block_hash[@title][second_key][true][f2_key] = {}
                      else
                        @block_hash[@title][final_key][true]["0 - 20"] = {} if value.eql?("≤ 20 Yr Term")
                        @block_hash[@title][final_key][true]["20 - Infinity"] = {} if value.eql?("> 20 Yr Term")
                        key = "0 - 20" if value.eql?("≤ 20 Yr Term")
                        key = "20 - Infinity" if value.eql?("> 20 Yr Term")
                      end
                    elsif index == 6 && rrr < 106 && value
                      another_key = value.eql?("≤ 85") ? set_range(value) : value
                      @block_hash[@title][final_key][true][key][another_key] = {} if another_key
                    end
                  elsif (107..112).to_a.include?(rrr) && index < 7 && value
                    if(rrr == 107) && (ccc == 4)
                      # for VLIP LPMI Adjustments
                        @title = "LPMI/PropertyType/FICO"
                        key = sheet_data.cell(rrr,ccc)
                        @block_hash[@title][true] = {}
                        @block_hash[@title][true][key] = {}
                    elsif (rrr == 109) && (ccc == 4)
                      @title = "LPMI/Term/FICO/LTV"
                      first_key  = sheet_data.cell(rrr,ccc)
                      second_key = sheet_data.cell(rrr,ccc + 1)
                      second_key.to_i
                      @block_hash[@title][true] = {}
                      @block_hash[@title][true][second_key.to_i.to_s] = {}
                    elsif (rrr > 108) && (ccc == 6)
                      key = get_value(value)
                      @block_hash[@title][true][second_key.to_i.to_s][key] = {} if second_key && key
                    end
                  end

                  if (115..118).to_a.include?(rrr) && ccc < 10
                    # for Subordinate Financing
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      key = get_value(key)
                      @block_hash[@title]["Subordinate Financing"][key] = {} unless @block_hash[@title].has_key?(key)
                    elsif index.eql?(7)
                      keyOfHash = sheet_data.cell(rrr,ccc)
                      keyOfHash = get_value(keyOfHash)
                      @block_hash[@title]["Subordinate Financing"][key][keyOfHash] = {}
                    end
                  end

                  if (120..121).to_a.include?(rrr)
                    # for Misc Adjusters
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      if key && key.eql?("Attached Condo > 75 LTV (>15yr Term)")
                        first_key = key.split(" >")[0].split(" ").last
                        @block_hash[@title][first_key] = {}
                        second_key = key.split(" ")[3] + ".01"
                        @block_hash[@title][first_key][second_key] = {}
                        third_key = key.split(" ")[5].split("(>")[1].split("yr")[0] + ".01"
                      end
                    end
                  end

                  if (rrr >= 116 && rrr <= 123) && index.eql?(15)
                    # for Loan Size Adjustments
                    loan_amount = get_value(sheet_data.cell(rrr, ccc))
                    @block_hash[@another_title][loan_amount] = {}
                  end

                  if [124,125].include?(rrr) && ccc == 6
                    # for Number Of Units
                    key = sheet_data.cell(rrr,ccc)
                    @block_hash[@title][key] = {}
                  end

                  # implementation of third key inside second key with value
                  if (81..88).to_a.include?(rrr) && index > 9 && value
                    #  for All Conforming
                    diff_of_row = rrr - 80
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 760") ? set_range(hash_key) : get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key][hash_key] = value unless @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 92 && rrr <= 105 && index >= 7 && value
                    # for Lender Paid MI Adjustments
                    diff_of_row = rrr - 92
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = set_range(hash_key).nil? ? get_value(hash_key) : set_range(hash_key)
                    if (93..95).to_a.include?(rrr)
                      if rrr.eql?(93)
                        @block_hash[@title][first_key][true][f1_key][hash_key] = value
                      elsif [94, 95].include?(rrr)
                        @block_hash[@title][second_key][true][f2_key][hash_key] = value
                      end
                    else
                      if rrr > 96
                        @block_hash[@title][final_key][true][key][another_key][hash_key] = value
                      end
                    end
                  end

                  if ((107..112).to_a.include?(rrr) && (ccc > 6))
                    # for VLIP LPMI Adjustments
                    diff_of_row = rrr - 92
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = get_value(hash_key)
                    if(rrr == 107)
                      # for VLIP LPMI Adjustments
                      @block_hash[@title][true][key][hash_key] = value if value && hash_key
                    elsif (109..112).to_a.include?(rrr)
                       @block_hash[@title][true][second_key.to_i.to_s][hash_key] = value if value && hash_key
                    end
                  elsif (115..118).to_a.include?(rrr) && ccc > 9 && ccc < 12 && value
                    # for Subordinate Financing
                    diff_of_row = rrr - 114
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 720") ? set_range(hash_key) : get_value(hash_key)
                    @block_hash[@title]["Subordinate Financing"][key][keyOfHash][hash_key] = value if hash_key.present?
                  end

                  if [120,121].include?(rrr) && ccc == 11
                    #for Misc Adjusters
                    if rrr.eql?(120)
                      @block_hash[@title][first_key][second_key][third_key] = value
                    else
                      first_key = sheet_data.cell(rrr,ccc - 5)
                      @block_hash["PropertyType"][first_key] = value if ("Manufacured Home").eql?(first_key)
                    end
                  end

                  if (116..123).to_a.include?(rrr) && ccc > 15 && value
                    #for Loan Size Adjustments
                    diff_of_row  = rrr - 115
                    loan_purpose = sheet_data.cell(rrr-diff_of_row,ccc)
                    @block_hash[@another_title][loan_amount][loan_purpose] = value
                  end

                  if [124,125].include?(rrr) && [9,10,11].include?(ccc)
                    # for Number Of Units
                    diff_of_row = rrr - 123
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = (hash_key.eql?("≤ 80") || hash_key.eql?("> 85")) ? set_range(hash_key) : get_value(hash_key)
                    @block_hash[@title][key][hash_key] = value if hash_key.present?
                  end

                  if rrr.eql?(127) && ccc.eql?(8)
                    @block_hash["LockDay"]["30"] = value
                    @block_hash["LockDay"]["45"] = value
                    @block_hash["LockDay"]["60"] = value
                  end
                end

                @allAdjustments[@title] = @block_hash[@title]
                if @another_title
                  @allAdjustments[@another_title] = @block_hash[@another_title]
                  @allAdjustments["PropertyType"] = @block_hash["PropertyType"]
                  @allAdjustments["LockDay"] = @block_hash["LockDay"]
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
    make_adjust(@allAdjustments, @sheet)
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def lp_open_acces_arms
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @programs_ids = []
    xlsx.sheets.each do |sheet|
      if (sheet == "LP Open Acces ARMs")
        sheet_data = xlsx.sheet(sheet)
        @adjustment_hash = {}
        @program_ids = []
        @fixed_data = []
        @sub_data = []
        @unit_data = []
        primary_key = ''
        secondry_key = ''
        misc_adj_key = ''
        term_key = ''
        ltv_key = ''
        cltv_key = ''
        misc_key = ''
        fixed_key = ''
        sub_data = ''
        key = ''
        @sheet = sheet
        (1..35).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              # title
              @title = sheet_data.cell(r,cc)

              # term
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # interest type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              # conforming
              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              # freddie_mac
              freddie_mac = false
              if @title.include?("Freddie Mac")
                freddie_mac = true
              end

              # fannie_mae
              fannie_mae = false
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                fannie_mae = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
               # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, sheet_name: sheet)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # Adjustments
        (37..71).each do |r|
          row = sheet_data.row(r)
          @fixed_data = sheet_data.row(39)
          @sub_data = sheet_data.row(47)
          @unit_data = sheet_data.row(56)
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

                # All LP Open Access ARMs
                if r >= 40 && r<= 45 && cc == 8# && cc <= 19 && cc != 15
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 40 && r<= 45 && cc > 8 && cc != 15 && cc <= 19
                  fixed_key = get_value @fixed_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = {}
                  @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = value
                end

                # Subordinate Financing Adjustments
                if r >= 48 && r <= 54 && cc == 5
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 48 && r <= 54 && cc == 6
                  cltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
                end
                if r >= 48 && r<= 54 && cc >= 9 && cc <= 10
                  sub_data = get_value @sub_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = {}
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = value
                end

                # Number Of Units Adjustments
                if r >= 57 && r <= 58 && cc == 3
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 57 && r <= 58 && cc > 3 && cc <= 7
                  sub_data = get_value @unit_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][ltv_key][sub_data] = {}
                  @adjustment_hash[primary_key][secondry_key][ltv_key][sub_data] = value
                end

                # Adjustments Applied after Cap
                if r >= 61 && r <= 67 && cc == 6
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 61 && r <= 67 && cc == 10
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = value
                end

                # Other Adjustments
                if r >= 69 && r <= 71 && cc == 3
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 69 && r <= 71 && cc == 10
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = value
                end
              end
            end
            (12..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if  value == "Misc Adjusters" || value == "Adjustment Caps"
                  key = value
                  @adjustment_hash[key] = {}
                end

                # Misc Adjustments
                if r >= 47 && r <= 57 && cc == 15
                  if value.include?("Condo")
                    misc_key = "Condo=>75.01=>15.01"
                  else
                    misc_key = value
                  end
                  @adjustment_hash[key][misc_key] = {}
                end
                if r >= 47 && r <= 57 && cc == 19
                  @adjustment_hash[key][misc_key] = value
                end

                # Adjustment Caps
                if r >= 62 && r <= 65 && cc == 16
                  misc_key = value
                  @adjustment_hash[key][misc_key] = {}
                end
                if r >= 62 && r <= 65 && cc == 17
                  term_key = get_value value
                  @adjustment_hash[key][misc_key][term_key] = {}
                end
                if r >= 62 && r <= 65 && cc == 18
                  ltv_key = get_value value
                  @adjustment_hash[key][misc_key][term_key][ltv_key] = {}
                end
                if r >= 62 && r <= 65 && cc == 19
                  @adjustment_hash[key][misc_key][term_key][ltv_key] = value
                end
                if r >= 67 && r <= 68 && cc == 12
                  misc_key = value
                  @adjustment_hash[key][misc_key] = {}
                end
                if r >= 67 && r <= 68 && cc == 16
                  @adjustment_hash[key][misc_key] = value
                end
              end
            end
          end
        end

        make_adjust(@adjustment_hash, @program_ids)
        create_program_association_with_adjustment(@sheet)
      end
    end
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def lp_open_access_105
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @programs_ids = []
    xlsx.sheets.each do |sheet|
      if (sheet == "LP Open Access_105")
        sheet_data = xlsx.sheet(sheet)
        @adjustment_hash = {}
        @program_ids = []
        @fixed_data = []
        @sub_data = []
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        term_key = ''
        caps_key = ''
        max_key = ''
        fixed_key = ''
        @sheet = sheet
        (1..61).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("LP Open Access 10yr Fixed >125 LTV"))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6
              # title
              @title = sheet_data.cell(r,cc)

              # term
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # interest type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              # conforming
              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              # freddie_mac
              freddie_mac = false
              if @title.include?("Freddie Mac")
                freddie_mac = true
              end

              # fannie_mae
              fannie_mae = false
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                fannie_mae = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
              # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, sheet_name: sheet)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # Adjustment
        (63..86).each do |r|
          row = sheet_data.row(r)
          @fixed_data = sheet_data.row(65)
          @sub_data = sheet_data.row(68)
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
                  secondry_key = "LoanSize/LoanType/Term/FICO/LTV"
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

                # All Fixed Conforming Adjustments
                if r == 66 && cc == 6
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r == 66 && cc > 6 && cc <= 19 && cc != 15
                  fixed_key = get_value @fixed_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = {}
                  @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = value
                end

                # Subordinate Financing
                if r == 69 && cc == 5
                  ltv_key = value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r == 69 && cc == 6
                  cltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
                end
                if r == 69 && cc >= 9 && cc <= 10
                  fixed_key = get_value @sub_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = {}
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = value
                end

                # Number Of Units
                if r >= 72 && r <= 73 && cc == 3
                  ltv_key = value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 72 && r <= 73 && cc == 5
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = value
                end

                # Adjustments Applied after Cap
                if r >= 76 && r <= 82 && cc == 6
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 76 && r <= 82 && cc == 10
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = value
                end

                # Other Adjustments
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

                # Misc Adjustments
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

                # Adjustment Caps
                if r > 76 && r <= 79 && cc == 16
                  caps_key = value
                  @adjustment_hash[primary_key][@key][caps_key] = {}
                end
                if r > 76 && r <= 79 && cc == 17
                  term_key = get_value value
                  @adjustment_hash[primary_key][@key][caps_key][term_key] = {}
                end
                if r > 76 && r <= 79 && cc == 18
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = {}
                end
                if r > 76 && r <= 79 && cc == 19
                  @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = value
                end

                # Other Adjustments
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
        make_adjust(@adjustment_hash, @program_ids)
        create_program_association_with_adjustment(@sheet)
      end
    end
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def jumbo_series_d
    @adjustment_hash = {}
    @property_hash = {}
    @state = {}
    primary_key = ''
    ltv_key = ''
    secondry_key = ''
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @programs_ids =[]
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_D")
        @sheet = sheet
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
                term =  program_heading[3]
                loan_type = program_heading[5]
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids  << @program.id
                 # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
                @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase", sheet_name: sheet)
                @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {} if key.present?
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    begin
                      @block_hash[main_key][key][15*c_i] = value if key.present? &&value.present?
                    rescue Exception => e
                    end
                  end
                  @data << value
                end
                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        #For Adjustments
        (41..71).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(44)
          if row.count >= 1
            (0..17).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "FICO/LTV Adjustments - Loan Amount ≤ $1MM"
                    @adjustment_hash["LoanAmount/FICO/LTV"] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1,000,000"] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1,000,000-Inf"] = {}
                  end
                  if value == "Feature Adjustments"
                    @property_hash["PropertyType/LTV"] = {}
                  end
                  if value == "State Adjustments"
                    @state["State"] = {}
                  end
                  # FICO/LTV Adjustments - Loan Amount ≤ $1MM
                  if r >= 45 && r <= 51 && cc == 3
                    if value.include?(">")
                      primary_key = value.tr('>= ','')+"-Inf"
                    else
                      primary_key = value
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1,000,000"][primary_key] = {}
                  end
                  if r >= 45 && r <= 51 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]  
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1,000,000"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1,000,000"][primary_key][ltv_key] = value
                  end
                  # State Adjustments
                  if r >= 45 && r <= 61 && cc == 11
                    secondry_key = value
                    @state["State"][secondry_key] = {}
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @state["State"][secondry_key] = new_val
                  end
                  if r >= 45 && r <= 61 && cc == 13
                    secondry_key = value
                    @state["State"][secondry_key] = {}
                    cc = cc + 2
                    new_val = sheet_data.cell(r,cc)
                    @state["State"][secondry_key] = new_val
                  end
                  if r >= 45 && r <= 61 && cc == 16
                    secondry_key = value
                    @state["State"][secondry_key] = {}
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @state["State"][secondry_key] = new_val
                  end
                  # FICO/LTV Adjustments - Loan Amount > $1MM
                  if r >= 55 && r <= 61 && cc == 3
                    if value.include?(">")
                      primary_key = value.tr('>= ','')+"-Inf"
                    else
                      primary_key = value
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1,000,000-Inf"][primary_key] = {}
                  end
                  if r >= 55 && r <= 61 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]  
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1,000,000-Inf"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1,000,000-Inf"][primary_key][ltv_key] = value
                  end
                  # Max Price
                  if r == 64 && cc == 11
                    @adjustment_hash["LoanType/Term"] = {}
                    @adjustment_hash["LoanType/Term"]["Fixed"] = {}
                    @adjustment_hash["LoanType/Term"]["Fixed"]["20"] = {}
                    @adjustment_hash["LoanType/Term"]["Fixed"]["30"] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @adjustment_hash["LoanType/Term"]["Fixed"]["20"] = new_val
                    @adjustment_hash["LoanType/Term"]["Fixed"]["30"] = new_val
                  end
                  if r == 65 && cc == 11
                    @adjustment_hash["LoanType/Term"]["Fixed"]["15"] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @adjustment_hash["LoanType/Term"]["Fixed"]["15"] = new_val
                  end
                  # Feature Adjustments
                  if r >= 65 && r <= 67 && cc == 2
                    if value == "Investment"
                      primary_key = "Investment Property"
                    else
                      primary_key = value
                    end
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 65 && r <= 67 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]  
                    end
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = {}
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = value
                  end
                  if r == 68 && cc == 2
                    @property_hash["RefinanceOption/LTV"] = {}
                    @property_hash["RefinanceOption/LTV"]["Cash Out"] = {}
                  end
                  if r == 68 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]  
                    end
                    @property_hash["RefinanceOption/LTV"]["Cash Out"][ltv_key] = {}
                    @property_hash["RefinanceOption/LTV"]["Cash Out"][ltv_key] = value
                  end
                  if r >= 69 && r <= 70 && cc == 2
                    primary_key = value
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 69 && r <= 70 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]  
                    end
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = {}
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = value
                  end
                  if r == 70 && cc == 2
                    @property_hash["MiscAdjuster/LTV"] = {}
                    @property_hash["MiscAdjuster/LTV"]["Escrow Waiver - except CA"] = {}
                  end
                  if r == 70 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]  
                    end
                    @property_hash["MiscAdjuster/LTV"]["Escrow Waiver - except CA"][ltv_key] = {}
                    @property_hash["MiscAdjuster/LTV"]["Escrow Waiver - except CA"][ltv_key] = value
                  end
                end
              rescue
                raise "value is nil at row = #{r} and column = #{cc}"
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@state]
        create_adjust(adjustment,sheet)
      end
    end
    # create adjustment for each program
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def lp_open_access
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @programs_ids = []
    xlsx.sheets.each do |sheet|
      if (sheet == "LP Open Access")
        sheet_data = xlsx.sheet(sheet)
        @adjustment_hash = {}
        @program_ids = []
        @fixed_data = []
        @sub_data = []
        @unit_data = []
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        unit_key = ''
        caps_key = ''
        term_key = ''
        max_key = ''
        fixed_key = ''
        sub_data = ''
        @sheet = sheet
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
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # rate type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              arm_basic = false
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              # conforming
              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              # freddie_mac
              freddie_mac = false
              if @title.include?("Freddie Mac")
                freddie_mac = true
              end

              # fannie_mae
              fannie_mae =false
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                fannie_mae = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
              # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, sheet_name: sheet)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        # Adjustment
        (63..97).each do |r|
          row = sheet_data.row(r)
          @fixed_data = sheet_data.row(65)
          @sub_data = sheet_data.row(73)
          @unit_data = sheet_data.row(82)
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
                  secondry_key = "LoanSize/LoanType/Term/FICO/LTV"
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
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 66 && r <= 71 && cc > 8 && cc <= 19 && cc != 15
                  fixed_key = @fixed_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = {}
                  @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = value
                end

                # Subordinate Adjustment
                if r >= 74 && r <= 80 && cc == 5
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 74 && r <= 80 && cc == 6
                  cltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
                end
                if r >= 74 && r <= 80 && cc >= 9 && cc <= 10
                  fixed_key = get_value @sub_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = {}
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = value
                end

                # Number of unit Adjustment
                if r >= 83 && r <= 84 && cc == 3
                  unit_key = value
                  @adjustment_hash[primary_key][secondry_key][unit_key] = {}
                end
                if r >= 83 && r <= 84 && cc > 3 && cc <= 7
                  fixed_key = get_value @unit_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][unit_key][fixed_key] = {}
                  @adjustment_hash[primary_key][secondry_key][unit_key][fixed_key] = value
                end

                # Loan Size Adjustments
                if r >= 87 && r <= 93 && cc == 6
                  ltv_key = get_value value
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
                  term_key = get_value value
                  @adjustment_hash[primary_key][@key][caps_key][term_key] = {}
                end
                if r > 86 && r <= 90 && cc == 18
                  ltv_key = get_value value
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
        make_adjust(@adjustment_hash, @program_ids)
        create_program_association_with_adjustment(@sheet)
      end
    end
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def jumbo_series_f
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_F")
        @adjustment_hash = {}
        @refinance_hash = {}
        @loan_amount = {}
        @state = {}
        @property_hash = {}
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        @sheet = sheet
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

              # term
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20/25/30 Yr")
                term = 20
              elsif @title.include?("10/15 Yr")
                term = 10
              end

              # rate type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              arm_basic = false
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 LIBOR ARM") || @title.include?("7/1 LIBOR ARM") || @title.include?("10/1 LIBOR ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: @loan_type,loan_purpose: "Purchase",arm_basic: arm_basic)
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # Adjustments
        (55..94).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(60)
          @cltv_data2 = sheet_data.row(59)
          @max_price_data = sheet_data.row(94)
          if row.compact.count >= 1
            (3..25).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
              if value.present?
                if value == "Purchase Transactions"
                  @adjustment_hash["LoanPurpose/FICO/LTV"] = {}
                  @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"] = {}
                  @state["State"] = {}
                end
                if value == "R/T Refinance Transactions"
                  @refinance_hash["RefinanceOption/FICO/LTV"] = {}
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"] = {}
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
                end
                if value == "Loan Amount Adjustments"
                  @loan_amount["LoanAmount/LTV"] = {}
                end
                if value == "Feature Adjustments"
                  @property_hash["PropertyType/LTV"] = {}
                end
                # Loan Amount Adjustments
                if r >= 60 && r <= 63 && cc == 15
                  if value.include?("≤")
                    ltv_key = "0-"+value.tr('A-Z≤ $ ','')+",000,000"
                  else
                    ltv_key = (value.tr('A-Z$ ','').split("-").first.to_f*1000000).to_s+"-"+(value.tr('A-Z$ ','').split("-").last.to_f*1000000).to_s
                  end
                  @loan_amount["LoanAmount/LTV"][ltv_key] = {}
                end
                if r >= 60 && r <= 63 && cc > 15 && cc <= 25
                  if @cltv_data2[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data2[cc-2]
                  end
                  @loan_amount["LoanAmount/LTV"][ltv_key][secondry_key] = {}
                  @loan_amount["LoanAmount/LTV"][ltv_key][secondry_key] = value
                end
                # Purchase Transactions Adjustment
                if r >= 61 && r <= 65 && cc == 3
                  if value.include?("≥")
                    primary_key = value.tr('≥ ','')+"-Inf"
                  else
                    primary_key = get_value value
                  end
                  @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key] = {}
                end
                if r >= 61 && r <= 65 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data[cc-2]
                  end
                  @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key][secondry_key] = {}
                  @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key][secondry_key] = value
                end
                # Feature Adjustments
                if r >= 68 && r <= 73 && cc == 15
                  primary_key = value
                  @property_hash["PropertyType/LTV"][primary_key] = {}
                end
                if r >= 68 && r <= 73 && cc > 15 && cc <= 25
                  if @cltv_data2[cc-2].present? && @cltv_data2[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data2[cc-2]
                  end
                  @property_hash["PropertyType/LTV"][primary_key][secondry_key] = {}
                  @property_hash["PropertyType/LTV"][primary_key][secondry_key] = value
                end
                # R/T Refinance Transactions Adjustment
                if r >= 69 && r <= 73 && cc == 3
                  if value.include?("≥")
                    primary_key = value.tr('≥ ','')+"-Inf"
                  else
                    primary_key = get_value value
                  end
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key] = {}
                end
                if r >= 69 && r <= 73 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data[cc-2]
                  end
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key][secondry_key] = {}
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key][secondry_key] = value
                end
                # # C/O Refinance Transactions Adjustment
                if r >= 77 && r <= 81 && cc == 3
                  if value.include?("≥")
                    primary_key = value.tr('≥ ','')+"-Inf"
                  else
                    primary_key = get_value value
                  end
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key] = {}
                end
                if r >= 77 && r <= 81 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data[cc-2]
                  end
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][secondry_key] = {}
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][secondry_key] = value
                end
                # State Adjustments
                if r == 86 && cc == 3
                  @state["State"]["FL"] = {}
                  @state["State"]["NV"] = {}
                end
                if r ==86 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data[cc-2]
                  end
                  @state["State"]["FL"][secondry_key] = {}
                  @state["State"]["NV"][secondry_key] = {}
                  @state["State"]["FL"][secondry_key] = value
                  @state["State"]["NV"][secondry_key] = value
                end
                if r == 87 && cc == 3
                  @state["State"]["CA"] = {}
                end
                if r ==87 && cc > 3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data[cc-2]
                  end
                  @state["State"]["CA"][secondry_key] = {}
                  @state["State"]["CA"][secondry_key] = value
                end
              end
              rescue
                raise "value is nil at row = #{r} and column = #{cc}"
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@refinance_hash,@loan_amount,@state,@property_hash]
        create_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def du_refi_plus_arms
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @programs_ids = []
    xlsx.sheets.each do |sheet|
      if (sheet == "Du Refi Plus ARMs")
        sheet_data = xlsx.sheet(sheet)
        @adjustment_hash = {}
        @program_ids = []
        @fixed_data = []
        @sub_data = []
        primary_key = ''
        secondry_key = ''
        fixed_key = ''
        ltv_key = ''
        cltv_key = ''
        sub_data = ''
        misc_key = ''
        adj_key = ''
        term_key = ''
        @sheet = sheet
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
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # rate type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              arm_basic = false
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              # conforming
              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              # freddie_mac
              freddie_mac = false
              if @title.include?("Freddie Mac")
                freddie_mac = true
              end

              # fannie_mae
              fannie_mae = false
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                fannie_mae = true
              end

              # High Balance
              if @title.include?("High Balance")
                jumbo_high_balance = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
              # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, sheet_name: sheet, jumbo_high_balance: jumbo_high_balance)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # Adjustments
        (37..70).each do |r|
          row = sheet_data.row(r)
          @fixed_data = sheet_data.row(39)
          @sub_data = sheet_data.row(49)
          if row.compact.count >= 1
            (3..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
                  primary_key = value
                  @adjustment_hash[primary_key] = {}
                end
                if value == "All DU Refi Plus Conforming ARMs (All Occupancies)" || value == "Subordinate Financing" || value == "Loan Size Adjustments"
                  secondry_key = value
                  @adjustment_hash[primary_key][secondry_key] = {}
                end

                # All du refi plus Adjustment
                if r >= 40 && r <= 47 && cc == 8
                  fixed_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][fixed_key] = {}
                end
                if r >= 40 && r <= 47 && cc >8 && cc <= 19
                  fixed_data = get_value @fixed_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][fixed_key][fixed_data] = value
                end

                # Subordinate Financing Adjustment
                if r >= 50 && r <= 54 && cc == 5
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 50 && r <= 54 && cc == 6
                  cltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
                end
                if r >= 50 && r <= 54 && cc > 6 && cc <= 10
                  sub_data = get_value @sub_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = value
                end

                # Other Adjustment
                if r >= 56 && r <= 57 && cc == 3
                  ltv_key = value
                  @adjustment_hash[primary_key][ltv_key] = {}
                end
                if r >= 56 && r <= 57 && cc == 8
                  @adjustment_hash[primary_key][ltv_key] = value
                end

                # Adjustments Applied after Cap
                if r >= 60 && r <= 66 && cc == 6
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 60 && r <= 66 && cc > 6 && cc <= 10
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = value
                end

                # Other Adjustment
                if r >= 69 && r <= 70 && cc == 3
                  ltv_key = value
                  @adjustment_hash[primary_key][ltv_key] = {}
                end
                if r >= 69 && r <= 70 && cc == 10
                  @adjustment_hash[primary_key][ltv_key] = value
                end
              end
            end
            (12..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)

              if value.present?
                if value == "Misc Adjusters" || value == "Adjustment Caps"
                  misc_key = value
                  @adjustment_hash[misc_key] = {}
                end

                # Misc Adjustments
                if r >= 49 && r <= 58 && cc == 15
                  if value.include?("Condo")
                    adj_key = "Condo/75"
                  else
                    adj_key = value
                  end
                  @adjustment_hash[misc_key][adj_key] = {}
                end
                if r >= 49 && r <= 58 && cc == 19
                  @adjustment_hash[misc_key][adj_key] = value
                end

                # Adjustment Caps
                if r >= 62 && r <= 64 && cc == 16
                  adj_key = value
                  @adjustment_hash[misc_key][adj_key] = {}
                end
                if r >= 62 && r <= 64 && cc == 17
                  term_key = get_value value
                  @adjustment_hash[misc_key][adj_key][term_key] = {}
                end
                if r >= 62 && r <= 64 && cc == 18
                  ltv_key = get_value value
                  @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = {}
                end
                if r >= 62 && r <= 64 && cc == 19
                  @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = value
                end
              end
            end
          end
        end
        make_adjust(@adjustment_hash, @program_ids)
        create_program_association_with_adjustment(@sheet)
      end
    end
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def jumbo_series_h
    @program_ids = []
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_H")
        sheet_data = xlsx.sheet(sheet)
        @sheet = sheet
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
                  # term
                  term = nil
                  program_heading = @title.split
                  if @title.include?("10yr") || @title.include?("10 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("15yr") || @title.include?("15 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("20yr") || @title.include?("20 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("25yr") || @title.include?("25 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("30yr") || @title.include?("30 Yr")
                    term = @title.scan(/\d+/)[0]
                  end

                  # rate type
                  if @title.include?("Fixed")
                    loan_type = "Fixed"
                  elsif @title.include?("ARM")
                    loan_type = "ARM"
                  elsif @title.include?("Floating")
                    loan_type = "Floating"
                  elsif @title.include?("Variable")
                    loan_type = "Variable"
                  else
                    loan_type = nil
                  end

                  # rate arm
                  arm_basic = false
                  if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 Yr ARM") || @title.include?("7/1 Yr ARM") || @title.include?("10/1 Yr ARM")
                    arm_basic = @title.scan(/\d+/)[0].to_i
                  end

                  # conforming
                  conforming = false
                  if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                    conforming = true
                  end

                  # freddie_mac
                  freddie_mac = false
                  if @title.include?("Freddie Mac")
                    freddie_mac = true
                  end

                  # fannie_mae
                  fannie_mae = false
                  if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                    fannie_mae = true
                  end

                  # High Balance
                  if @title.include?("High Balance")
                    jumbo_high_balance = true
                  end

                  # Purchase & Refinance
                  if @title.include?("Purchase")
                    loan_purpose = "Purchase"
                  elsif @title.include?("Refinance")
                    loan_purpose = "Refinance"
                  end

                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @program_ids << @program.id
                # Loan Limit Type
                if @title.include?("Non-Conforming")
                  @program.loan_limit_type << "Non-Conforming"
                end
                if @title.include?("Conforming")
                  @program.loan_limit_type << "Conforming"
                end
                if @title.include?("Jumbo")
                  @program.loan_limit_type << "Jumbo"
                end
                if @title.include?("High Balance")
                  @program.loan_limit_type << "High Balance"
                end
                @program.save
                @program.update(term: term,loan_type: loan_type,loan_purpose: loan_purpose ,arm_basic: arm_basic )
                @program.adjustments.destroy_all

                @block_hash = {}
                key = ''
                main_key = ''
              if @program.term.present?
                main_key = loan_purpose.to_s + "/" +"Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*c_i
                        @program.save
                      end
                      @block_hash[main_key][key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end

        #For Adjustments
        xlsx.sheet(sheet).each_with_index do |sheet_row, index|
          index = index+ 1
          if sheet_row.include?("Jumbo Series H - Adjustments")
            (index..xlsx.sheet(sheet).last_row).each do |adj_row|
              # First Adjustment
              if xlsx.sheet(sheet).row(adj_row).include?("State Adjustments")
                begin
                  key_array = ""
                  rr = adj_row
                  cc = 12
                  @state_hash = {}
                  main_key = "State"
                  @state_hash[main_key] = {}
                  @right_adj = {}
                  key = ''
                  (1..11).each do |max_row|
                    column_count = 1
                    rrr = rr + max_row
                    row = xlsx.sheet(sheet).row(rrr)
                    if row.include?("State")
                      key_array = row.compact[5..12]
                    end
                    (0..8).each do |max_column|
                      ccc = cc + max_column
                      value = xlsx.sheet(sheet).cell(rrr,ccc)
                      if !(row.include?("State"))
                        if (ccc == 12)
                          key = value
                          @state_hash[main_key][key] = {}
                        else
                          if value != nil
                            @state_hash[main_key][key][key_array[column_count]] = value
                            column_count = column_count + 1
                          end
                        end
                      end
                    end
                  end
                  make_adjust(@state_hash, @program_ids)
                rescue => e
                end
              end

              # Second Adjustment
              if xlsx.sheet(sheet).row(adj_row).include?("Credit Score")
                begin
                  key_array = []
                  rr = adj_row
                  cc = 12
                  @credit_score = {}
                  main_key = "Credit Score"
                  @credit_score[main_key] = {}
                  (1..7).each do |max_row|
                    column_count = 0
                    rrr = rr + max_row
                    row = xlsx.sheet(sheet).row(rrr)

                    if row.include?("CLTV -->")
                      row.compact[5..9].each do |row_val|
                        val = row_val.split
                        if val.include?("≤") && !val.include?("CLTV")
                          key_array << 0
                        elsif !val.include?("CLTV")
                          key_array << row_val.split("-")[0].to_i.round
                        end
                      end
                    end

                    (0..5).each do |max_column|
                      ccc = cc + max_column
                      value = xlsx.sheet(sheet).cell(rrr,ccc)
                      if !row.include?("CLTV -->")
                        if ccc == 12
                          key = value.split("-")[0]
                          @credit_score[main_key][key] = {}
                        else
                          @credit_score[main_key][key][key_array[column_count]] = value if value != nil
                          column_count = column_count + 1 if value != nil
                        end
                      end
                    end
                  end
                  make_adjust(@credit_score, @program_ids)
                rescue Exception => e
                end
              end

              # Third Adjustment
              if xlsx.sheet(sheet).row(adj_row).include?("Other Adjustments")
                begin
                  rr = adj_row
                  cc = 13

                  (1..2).each do |max_row|
                    rrr = rr + max_row
                    row = xlsx.sheet(sheet).row(rrr)

                    if row.include?("Loan Amount >=$1MM")
                      loan_amount = {}
                      main_key = "LoanAmount"
                      loan_amount[main_key] = {}
                      key = 1000000 if xlsx.sheet(sheet).cell(rrr,cc).split[2].include?(">")
                      value = xlsx.sheet(sheet).cell(rrr,cc+4)

                      loan_amount[main_key][key] = {}
                      loan_amount[main_key][key] = value
                      make_adjust(loan_amount, @program_ids)
                    end

                    if row.include?("Second Home")
                      second_home = {}
                      main_key = "PropertyType"
                      second_home[main_key] = {}

                      key = "2nd Home" if xlsx.sheet(sheet).cell(rrr,cc).include?("Second Home")
                      value = xlsx.sheet(sheet).cell(rrr,cc+4)

                      second_home[main_key][key] = {}
                      second_home[main_key][key] = value
                      make_adjust(second_home, @program_ids)
                    end
                  end
                rescue Exception => e
                end
              end


              # Fourth Adjustment
              if xlsx.sheet(sheet).row(adj_row).include?("Cash Out Refi")
                if adj_row == 95
                  begin
                    rr = adj_row
                    cc = 15
                    @data_hash = {}
                    main_key = "LoanPurpose/RefinanceOption/LTV"
                    key = "True"
                    key1 = "Cash Out"
                    @data_hash[main_key] = {}
                    @data_hash[main_key][key] = {}
                    @data_hash[main_key][key][key1] = {}

                    (0..2).each do |max_row|
                      rrr = rr + max_row
                      row = xlsx.sheet(sheet).row(rrr)
                      cell_value = xlsx.sheet(sheet).cell(rrr,cc)

                      key2 = get_value(cell_value)
                      value = xlsx.sheet(sheet).cell(rrr,cc+2)

                      @data_hash[main_key][key][key1][key2] = {}
                      @data_hash[main_key][key][key1][key2] = value

                    end
                    make_adjust(@data_hash, @program_ids)
                  rescue Exception => e
                  end
                end
              end

            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def du_refi_plus_fixed_rate_105
    @program_ids = []
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Du Refi Plus Fixed Rate_105")
        sheet_data = xlsx.sheet(sheet)
        @sheet = sheet
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
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # rate type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              # conforming
              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              # freddie_mac
              freddie_mac = false
              if @title.include?("Freddie Mac")
                freddie_mac = true
              end

              # fannie_mae
              fannie_mae = false
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                fannie_mae = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
              # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, sheet_name: sheet)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        #For Adjustments
        xlsx.sheet(sheet).each_with_index do |sheet_row, index|
          index = index+ 1
          if sheet_row.include?("Loan Level Price Adjustments: See Adjustment Caps")
            (index..xlsx.sheet(sheet).last_row).each do |adj_row|
              # First Adjustment
              if adj_row == 65
                begin
                  key = ''
                  key_array = []
                  rr = adj_row
                  cc = 3
                  @occupancy_hash = {}
                  main_key = "All Occupancies"
                  @occupancy_hash[main_key] = {}

                  (0..2).each do |max_row|
                    column_count = 0
                    rrr = rr + max_row
                    row = xlsx.sheet(sheet).row(rrr)

                    if rrr == rr
                      row.compact.each do |row_val|
                        val = row_val.split
                        if val.include?("<")
                          key_array << 0
                        else
                          key_array << row_val.split("-")[0].to_i.round if row_val.include?("-")
                          key_array << row_val.split[1].to_i.round if row_val.include?(">")
                        end
                      end
                    end

                    (0..16).each do |max_column|
                      ccc = cc + max_column
                      value = xlsx.sheet(sheet).cell(rrr,ccc)
                      if row.include?("All Occupancies > 15 Yr Terms")
                        if value != nil && value.to_s.include?(">") && value != "All Occupancies > 15 Yr Terms" && !value.is_a?(Numeric)
                          key = value.gsub(/[^0-9A-Za-z]/, '')
                          @occupancy_hash[main_key][key] = {}
                        elsif (value != nil) && !value.is_a?(String)
                          @occupancy_hash[main_key][key][key_array[column_count]] = value
                          column_count = column_count + 1
                        end
                      end
                    end
                  end
                  make_adjust(@occupancy_hash, @program_ids)
                rescue Exception => e
                end
              end

              # Second Adjustment(Adjustment Caps)
              if adj_row == 86
                begin
                  key_array = ""
                  rr = adj_row
                  cc = 16
                  @adjustment_cap = {}
                  main_key = "Adjustment Caps"
                  @adjustment_cap[main_key] = {}
                  key = ''

                  (0..4).each do |max_row|
                    column_count = 1
                    rrr = rr + max_row
                    row = xlsx.sheet(sheet).row(rrr)
                    if rrr == 86
                      key_array = row.compact
                    end

                    (0..3).each do |max_column|
                      ccc = cc + max_column
                      value = xlsx.sheet(sheet).cell(rrr,ccc)
                      if ccc == 16
                        key = value if value != nil
                        @adjustment_cap[main_key][key] = {} if value != nil
                      else
                        if !key_array.include?(value)
                          @adjustment_cap[main_key][key][key_array[column_count]] = value if value != nil
                          column_count = column_count + 1 if value != nil
                        end
                      end
                    end
                  end
                  make_adjust(@adjustment_cap, @program_ids)
                rescue Exception => e
                end
              end

              # Third Adjustment
              if xlsx.sheet(sheet).row(adj_row).include?("Max YSP")
                begin
                  rr = adj_row
                  cc = 4
                  @max_ysp_hash = {}
                  main_key = "Max YSP"
                  @max_ysp_hash[main_key] = {}
                  row = xlsx.sheet(sheet).row(rr)
                  @max_ysp_hash[main_key] = row.compact[5]
                  make_adjust(@max_ysp_hash, @program_ids)
                rescue Exception => e
                end
              end

              # Fourth Adjustment (Adjustments Applied after Cap)
              if xlsx.sheet(sheet).row(adj_row).include?("Loan Size Adjustments")
                begin
                  rr = adj_row
                  cc = 6
                  @loan_size = {}
                  main_key = "Loan Size / Loan Type"
                  @loan_size[main_key] = {}

                  (0..6).each do |max_row|
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
                      raise "value is nil at row = #{rrr} and column = #{ccc}" unless value || key
                      @loan_size[main_key][key] = value
                    end
                  end
                  make_adjust(@loan_size, @program_ids)
                rescue => e
                end
              end

              # Fifth Adjustment(Misc Adjusters)
              if xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
                begin
                  rr = adj_row
                  cc = 15
                  @cando_hash = {}
                  main_key = "PropertyType/LTV/Term"
                  @cando_hash[main_key] = {}

                  (0..6).each do |max_row|
                    @data = []
                    rrr = rr + max_row
                    ccc = cc
                    key = xlsx.sheet(sheet).cell(rrr,ccc)

                    if key.include?("Condo")
                      val = key.split
                      key1 = "Condo"
                      key2 = val[1].gsub(/[^0-9A-Za-z]/, '')
                      key3 = val[3].gsub(/[^0-9A-Za-z]/, '').split("yr")[0]
                      value = xlsx.sheet(sheet).cell(rrr,ccc+4)
                      @cando_hash[main_key][key1] = {}
                      @cando_hash[main_key][key1][key2] = {}
                      @cando_hash[main_key][key1][key2][key3] = value
                    end

                    if key == "Manufactured Home"
                      key1 = "Manufactured Home"
                      key2 = 0
                      value = xlsx.sheet(sheet).cell(rrr,ccc+4)
                      @cando_hash[main_key][key1] = {}
                      @cando_hash[main_key][key1][key2] = {}
                      @cando_hash[main_key][key1][key2] = value
                    end
                  end
                  make_adjust(@cando_hash, @program_ids)
                rescue => e
                end
              end

              # Sixth Adjustment(Misc Adjusters (2-4 Units))
              if xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
                begin
                  rr = adj_row
                  cc = 15
                  @unit_hash = {}
                  main_key = "PropertyType/LTV"
                  @unit_hash[main_key] = {}

                  rrr = rr + 1
                  ccc = cc
                  key = xlsx.sheet(sheet).cell(rrr,ccc)

                  if key.include?("Units")
                    key1 = "2-4 unit"
                    value = xlsx.sheet(sheet).cell(rrr,ccc+4)
                    @unit_hash[main_key][key1] = {}
                    @unit_hash[main_key][key1] = value
                  end
                  make_adjust(@unit_hash, @program_ids)
                rescue Exception => e
                end
              end


              # Seventh Adjustment(Misc Adjusters)
              if xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
                begin
                  rr = adj_row
                  cc = 15
                  @data_hash = {}
                  main_key = "MiscAdjuster"
                  @data_hash[main_key] = {}

                  (0..2).each do |max_row|
                    rrr = rr + max_row
                    ccc = cc
                    key = xlsx.sheet(sheet).cell(rrr,ccc)

                    if !key.include?("Units")
                      key1 = key.include?(">") ? key.split(" >")[0] : key
                      value = xlsx.sheet(sheet).cell(rrr,ccc+4)
                      @data_hash[main_key][key1] = {}
                      @data_hash[main_key][key1] = value
                    end
                  end
                  make_adjust(@data_hash, @program_ids)
                rescue Exception => e
                end
              end

              # LTV Adjustment(Misc Adjusters)
              if xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
                begin
                  rr = adj_row
                  cc = 15
                  @ltv_hash = {}
                  main_key = "LTV"
                  @ltv_hash[main_key] = {}


                  (0..6).each do |max_row|
                    rrr = rr + max_row
                    ccc = cc
                    key = xlsx.sheet(sheet).cell(rrr,ccc)

                    if key.include?("LTV") && !key.include?("Condo")
                      key1 = key.split[1].to_i.round
                      key2 = key.include?("<") ? 0 : 30
                      value = xlsx.sheet(sheet).cell(rrr,ccc+4)
                      @ltv_hash[main_key][key1] = {} if @ltv_hash[main_key] == {}
                      @ltv_hash[main_key][key1][key2] = {}
                      @ltv_hash[main_key][key1][key2] = value
                    end
                  end
                  make_adjust(@ltv_hash, @program_ids)
                rescue Exception => e
                end
              end

              # CA Escrow Waiver Adjustment
              if xlsx.sheet(sheet).row(adj_row).include?("Expanded Approval **")
                begin
                  rr = adj_row
                  cc = 3
                  @misc_adjuster = {}
                  main_key = "MiscAdjuster"
                  @misc_adjuster[main_key] = {}

                  (0..2).each do |max_row|
                    rrr = rr + max_row
                    ccc = cc
                    key = xlsx.sheet(sheet).cell(rrr,ccc)

                    if key.include?("CA Escrow Waiver") || key.include?("Expanded Approval **")
                      value = xlsx.sheet(sheet).cell(rrr,ccc+7)
                      @misc_adjuster[main_key][key] = {}
                      @misc_adjuster[main_key][key] = value
                    end
                  end
                  make_adjust(@misc_adjuster, @program_ids)
                rescue Exception => e
                end
              end

              # Subordinate Financing Adjustment
              if xlsx.sheet(sheet).row(adj_row).include?("Subordinate Financing")
                begin
                  rr = adj_row
                  cc = 6
                  @subordinate_hash = {}
                  main_key = "FinancingType/LTV/CLTV/FICO"
                  key1 = "Subordinate Financing"

                  sub_key1 = row.compact[2].include?("<") ? 0 : row.compact[2].split(" ")[1].to_i
                  sub_key2 = row.compact[3].include?(">") ? row.compact[3].split(" ")[1].to_i : row.compact[3].to_i

                  @subordinate_hash[main_key] = {}
                  @subordinate_hash[main_key][key1] = {}

                  (1..2).each do |max_row|
                    rrr = rr + max_row
                    ccc = cc
                    key = xlsx.sheet(sheet).cell(rrr,ccc)

                    if key.include?(">") || key == "ALL"
                      key2 = (key.include?(">")) ? key.gsub(/[^0-9A-Za-z]/, '') : key
                      value = xlsx.sheet(sheet).cell(rrr,ccc+3)
                      value1 = xlsx.sheet(sheet).cell(rrr,ccc+4)

                      @subordinate_hash[main_key][key1][key2] ={}
                      @subordinate_hash[main_key][key1][key2][sub_key1] = value
                      @subordinate_hash[main_key][key1][key2][sub_key2] = value1
                    end
                  end
                  make_adjust(@subordinate_hash, @program_ids)
                rescue => e
                end
              end
            end
          end
        end
      end
    end
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def jumbo_series_i
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @programs_ids = []
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_I")
        sheet_data = xlsx.sheet(sheet)
        @adjustment_hash = {}
        @program_ids = []
        @credit_data = []
        primary_key = ''
        key = ''
        cltv_key = ''
        key1 = ''
        cltv_key1 = ''
        credit_data = ''
        main_key = ''
        @sheet = sheet
        # programs
        (2..32).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8)/ (r == 21)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              @title = sheet_data.cell(r,cc)
              if @title.present?
                program_heading = @title.split
                # term
                  term = nil
                  program_heading = @title.split
                  if @title.include?("10yr") || @title.include?("10 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("15yr") || @title.include?("15 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("20yr") || @title.include?("20 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("25yr") || @title.include?("25 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("30yr") || @title.include?("30 Yr")
                    term = @title.scan(/\d+/)[0]
                  end

                  # rate type
                  if @title.include?("Fixed")
                    loan_type = "Fixed"
                  elsif @title.include?("ARM")
                    loan_type = "ARM"
                  elsif @title.include?("Floating")
                    loan_type = "Floating"
                  elsif @title.include?("Variable")
                    loan_type = "Variable"
                  else
                    loan_type = nil
                  end

                  # rate arm
                  if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                    arm_basic = @title.scan(/\d+/)[0].to_i
                  end


                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @program_ids << @program.id
                  # Loan Limit Type
                if @title.include?("Non-Conforming")
                  @program.loan_limit_type << "Non-Conforming"
                end
                if @title.include?("Conforming")
                  @program.loan_limit_type << "Conforming"
                end
                if @title.include?("Jumbo")
                  @program.loan_limit_type << "Jumbo"
                end
                if @title.include?("High Balance")
                  @program.loan_limit_type << "High Balance"
                end
                @program.save
                @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase" ,arm_basic: arm_basic, sheet_name: sheet )
                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                new_key = ''
                if @program.term.present?
                  new_key = "Term/LoanType/InterestRate/LockPeriod"
                else
                  new_key = "InterestRate/LockPeriod"
                end
                @block_hash[new_key] = {}
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[new_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*c_i
                        @program.save
                      end
                      @block_hash[new_key][key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end


        key = ''
        another_key = ''
        keyOfHash = ''
        # for Misc Adjusters
        first_key   = ''
        second_key  = ''
        third_key   = ''
        final_key   = ''
        f1_key      = ''
        f2_key      = ''
        loan_amount = ''
        # Adjustments
        (40..73).each do |row|
          unless @adjustment_hash.has_key?("LoanType/FICO/LTV")
            #  for Credit Score table
            @adjustment_hash["LoanType/FICO/LTV"] = {}
            @adjustment_hash["LoanType/FICO/LTV"]["Fixed"] = {}
            @adjustment_hash["LoanType/FICO/LTV"]["ARM"] = {}
          end

          unless @adjustment_hash.has_key?("LoanType/LoanAmount/LTV")
            # for Loan Amount table
            @adjustment_hash["LoanType/LoanAmount/LTV"] = {}
            @adjustment_hash["LoanType/LoanAmount/LTV"]["Fixed"] = {}
            @adjustment_hash["LoanType/LoanAmount/LTV"]["ARM"] = {}
          end

          unless @adjustment_hash.has_key?("LoanType/PropertyType/LTV")
            # for Purpose/Property Type table
            @adjustment_hash["LoanType/PropertyType/LTV"] = {}
            @adjustment_hash["LoanType/PropertyType/LTV"]["Fixed"] = {}
            @adjustment_hash["LoanType/PropertyType/LTV"]["ARM"] = {}
          end

          unless @adjustment_hash.has_key?("LoanType/RefinanceOption/Term/LTV")
            @adjustment_hash["LoanType/RefinanceOption/Term/LTV"] = {}
            @adjustment_hash["LoanType/RefinanceOption/Term/LTV"]["Fixed"] = {}
            @adjustment_hash["LoanType/RefinanceOption/Term/LTV"]["ARM"] = {}
          end

          unless @adjustment_hash.has_key?("LoanType/RefinanceOption/LTV")
            @adjustment_hash["LoanType/RefinanceOption/LTV"] = {}
            @adjustment_hash["LoanType/RefinanceOption/LTV"]["Fixed"] = {}
            @adjustment_hash["LoanType/RefinanceOption/LTV"]["ARM"] = {}
          end

          (3..19).each do |column|
            value = sheet_data.cell(row,column)

            # prepare first key inside hash
            if((row >= 41 && row <= 45) && column.eql?(3))
              first_key = set_range(value) || get_value(value)
              @adjustment_hash["LoanType/FICO/LTV"]["Fixed"][first_key] = {}
            end

            if((row >= 50 && row <= 53) && column.eql?(3))
              first_key = set_range(value) || get_value(value)
              @adjustment_hash["LoanType/LoanAmount/LTV"]["Fixed"][first_key] = {}
            end

            if((row >= 58 && row <= 62) && column.eql?(3))
            end

            if((row >= 41 && row <= 45) && column.eql?(12))
            end

            if((row >= 50 && row <= 53) && column.eql?(12))
            end

            if((row >= 58 && row <= 62) && column.eql?(12))
            end
          end
        end
        make_adjust(@adjustment_hash, @program_ids)
        create_program_association_with_adjustment(@sheet)
      end
    end
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def du_refi_plus_fixed_rate
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @programs_ids = []
    xlsx.sheets.each do |sheet|
      if (sheet == "Du Refi Plus Fixed Rate")
        sheet_data = xlsx.sheet(sheet)
        @adjustment_hash = {}
        @program_ids = []
        @fixed_data = []
        @sub_data = []
        sub_data = ''
        primary_key = ''
        secondry_key = ''
        fixed_key = ''
        ltv_key = ''
        cltv_key = ''
        misc_key = ''
        adj_key = ''
        term_key = ''
        @sheet = sheet
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
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # rate type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              # conforming
              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              # High Balance
              jumbo_high_balance = false
              if @title.include?("High Balance")
                jumbo_high_balance = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
              # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming, arm_basic: arm_basic, sheet_name: sheet, jumbo_high_balance: jumbo_high_balance)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # Adjustments
        (63..94).each do |r|
          row = sheet_data.row(r)
          @fixed_data = sheet_data.row(65)
          @sub_data = sheet_data.row(75)
          if row.compact.count >= 1
            (3..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
                  primary_key = value
                  @adjustment_hash[primary_key] = {}
                end
                if (r == 65 && cc == 3)
                  secondry_key = "LoanSize/LoanType/Term/FICO/LTV"
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                if value == "Subordinate Financing"
                  secondry_key = "FinancingType/LTV/CLTV/FICO"
                  @adjustment_hash[primary_key][secondry_key] = {}
                end
                if value == "Loan Size Adjustments"
                  secondry_key = "Loan Size Adjustments"
                  @adjustment_hash[primary_key][secondry_key] = {}
                end

                # All Fixed Confoming Adjustment
                if r >= 66 && r <= 73 && cc == 8
                  fixed_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][fixed_key] = {}
                end
                if r >= 66 && r <= 73 && cc >8 && cc <= 19
                  fixed_data = get_value @fixed_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][fixed_key][fixed_data] = value
                end

                # Subordinate Financing Adjustment
                if r >= 76 && r <= 80 && cc == 5
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 76 && r <= 80 && cc == 6
                  cltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
                end
                if r >= 76 && r <= 80 && cc > 6 && cc <= 10
                  sub_data = get_value @sub_data[cc-2]
                  @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = value
                end

                # Adjustments Applied after Cap
                if r >= 83 && r <= 89 && cc == 6
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
                end
                if r >= 83 && r <= 89 && cc > 6 && cc <= 10
                  @adjustment_hash[primary_key][secondry_key][ltv_key] = value
                end

                # Other Adjustment
                if r >= 92 && r <= 94 && cc == 3
                  ltv_key = value
                  @adjustment_hash[primary_key][ltv_key] = {}
                end
                if r >= 92 && r <= 94 && cc == 10
                  @adjustment_hash[primary_key][ltv_key] = value
                end
              end
            end
            (12..19).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)

              if value.present?
                if value == "Misc Adjusters" || value == "Adjustment Caps"
                  misc_key = value
                  @adjustment_hash[misc_key] = {}
                end

                # Misc Adjustments
                if r >= 75 && r <= 83 && cc == 15
                  if value.include?("Condo")
                    adj_key = "Condo/75/15"
                  else
                    adj_key = value
                  end
                  @adjustment_hash[misc_key][adj_key] = {}
                end
                if r >= 75 && r <= 83 && cc == 19
                  @adjustment_hash[misc_key][adj_key] = value
                end

                # Other Adjustments
                if r == 85 && cc == 13
                  adj_key = value
                  @adjustment_hash[adj_key] = {}
                end
                if r == 85 && cc == 17
                  @adjustment_hash[adj_key] = value
                end

                # Adjustment Caps
                if r >= 89 && r <= 93 && cc == 16
                  adj_key = value
                  @adjustment_hash[misc_key][adj_key] = {}
                end
                if r >= 89 && r <= 93 && cc == 17
                  term_key = get_value value
                  @adjustment_hash[misc_key][adj_key][term_key] = {}
                end
                if r >= 89 && r <= 93 && cc == 18
                  ltv_key = get_value value
                  @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = {}
                end
                if r >= 89 && r <= 93 && cc == 19
                  @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = value
                end
              end
            end
          end
        end
        make_adjust(@adjustment_hash, @program_ids)
        create_program_association_with_adjustment(@sheet)
      end
    end
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def jumbo_series_jqm
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @programs_ids = []
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_JQM")
        sheet_data = xlsx.sheet(sheet)
        @program_ids = []
        @adjustment_hash = {}
        @refinance_hash = {}
        @loan_amount = {}
        @state = {}
        @property_hash = {}
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        @sheet = sheet
        (2..60).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8)/ (r == 21)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 6 + max_column*6 # (6 / 12 / 18)
              @title = sheet_data.cell(r,cc)
              if @title.present?
                program_heading = @title.split
                # term
                term = nil
                program_heading = @title.split
                if @title.include?("10yr") || @title.include?("10 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("15yr") || @title.include?("15 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("20yr") || @title.include?("20 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("25yr") || @title.include?("25 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("30yr") || @title.include?("30 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("20/25/30 Yr")
                  term = 20
                elsif @title.include?("10/15 Yr")
                  term = 10
                end
                # rate type
                if @title.include?("Fixed")
                  loan_type = "Fixed"
                elsif @title.include?("ARM")
                  loan_type = "ARM"
                elsif @title.include?("Floating")
                  loan_type = "Floating"
                elsif @title.include?("Variable")
                  loan_type = "Variable"
                else
                  loan_type = nil
                end

                # rate arm
                arm_basic = false
                if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 LIBOR ARM") || @title.include?("7/1 LIBOR ARM") || @title.include?("10/1 LIBOR ARM")
                  arm_basic = @title.scan(/\d+/)[0].to_i
                end
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @program_ids << @program.id
                # Loan Limit Type
                if @title.include?("Non-Conforming")
                  @program.loan_limit_type << "Non-Conforming"
                end
                if @title.include?("Conforming")
                  @program.loan_limit_type << "Conforming"
                end
                if @title.include?("Jumbo")
                  @program.loan_limit_type << "Jumbo"
                end
                if @title.include?("High Balance")
                  @program.loan_limit_type << "High Balance"
                end
                @program.save
                @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase" ,arm_basic: arm_basic, sheet_name: sheet )
                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                main_key = ''
                if @program.term.present?
                  main_key = "Term/LoanType/InterestRate/LockPeriod"
                else
                  main_key = "InterestRate/LockPeriod"
                end
                @block_hash[main_key] = {}
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*c_i
                        @program.save
                      end
                      @block_hash[main_key][key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break #terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end

        # Adjustments
        (64..99).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(67)
          @cltv_data2 = sheet_data.row(66)
          @max_price_data = sheet_data.row(94)
          if row.compact.count >= 1
            (3..25).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
              if value.present?
                if value == "Purchase Transactions"
                  @adjustment_hash["LoanPurpose/FICO/LTV"] = {}
                  @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"] = {}
                  @state["State"] = {}
                end
                if value == "R/T Refinance Transactions"
                  @refinance_hash["RefinanceOption/FICO/LTV"] = {}
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"] = {}
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
                end
                if value == "Loan Amount Adjustments"
                  @loan_amount["LoanAmount/LTV"] = {}
                end
                if value == "Feature Adjustments"
                  @property_hash["PropertyType/LTV"] = {}
                end
                # Loan Amount Adjustments
                if r >= 67 && r <= 70 && cc == 15
                  if value.include?("≤")
                    ltv_key = "0-"+value.tr('A-Z≤ $ ','')+",000,000"
                  else
                    ltv_key = (value.tr('A-Z$ ','').split("-").first.to_f*1000000).to_s+"-"+(value.tr('A-Z$ ','').split("-").last.to_f*1000000).to_s
                  end
                  @loan_amount["LoanAmount/LTV"][ltv_key] = {}
                end
                if r >= 67 && r <= 70 && cc > 15 && cc <= 25
                  if @cltv_data2[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data2[cc-2]
                  end
                  @loan_amount["LoanAmount/LTV"][ltv_key][secondry_key] = {}
                  @loan_amount["LoanAmount/LTV"][ltv_key][secondry_key] = value
                end
                # Purchase Transactions Adjustment
                if r >= 68 && r <= 74 && cc == 3
                  if value.include?("≥")
                    primary_key = value.tr('≥ ','')+"-Inf"
                  else
                    primary_key = get_value value
                  end
                  @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key] = {}
                end
                if r >= 68 && r <= 74 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data[cc-2]
                  end
                  @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key][secondry_key] = {}
                  @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key][secondry_key] = value
                end
                # Feature Adjustments
                if r >= 75 && r <= 80 && cc == 15
                  primary_key = value
                  @property_hash["PropertyType/LTV"][primary_key] = {}
                end
                if r >= 75 && r <= 80 && cc > 15 && cc <= 25
                  if @cltv_data2[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data2[cc-2]
                  end
                  @property_hash["PropertyType/LTV"][primary_key][secondry_key] = {}
                  @property_hash["PropertyType/LTV"][primary_key][secondry_key] = value
                end
                # R/T Refinance Transactions Adjustment
                if r >= 78 && r <= 84 && cc == 3
                  if value.include?("≥")
                    primary_key = value.tr('≥ ','')+"-Inf"
                  else
                    primary_key = get_value value
                  end
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key] = {}
                end
                if r >= 78 && r <= 84 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data[cc-2]
                  end
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key][secondry_key] = {}
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key][secondry_key] = value
                end
                # # C/O Refinance Transactions Adjustment
                if r >= 88 && r <= 94 && cc == 3
                  if value.include?("≥")
                    primary_key = value.tr('≥ ','')+"-Inf"
                  else
                    primary_key = get_value value
                  end
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key] = {}
                end
                if r >= 88 && r <= 94 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data[cc-2]
                  end
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][secondry_key] = {}
                  @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][secondry_key] = value
                end
                # State Adjustments
                if r == 99 && cc == 3
                  @state["State"]["FL"] = {}
                  @state["State"]["NV"] = {}
                end
                if r ==99 && cc >3 && cc <= 13
                  if @cltv_data[cc-2].include?("≤")
                    secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                  else
                    secondry_key = get_value @cltv_data[cc-2]
                  end
                  @state["State"]["FL"][secondry_key] = {}
                  @state["State"]["NV"][secondry_key] = {}
                  @state["State"]["FL"][secondry_key] = value
                  @state["State"]["NV"][secondry_key] = value
                end
              end
              rescue
                raise "value is nil at row = #{r} and column = #{cc}"
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@refinance_hash,@loan_amount,@state,@property_hash]
        create_adjust(adjustment,sheet)
        create_program_association_with_adjustment(@sheet)
      end
    end
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def dream_big
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @programs_ids = []
    xlsx.sheets.each do |sheet|
      if (sheet == "Dream Big")
        sheet_data = xlsx.sheet(sheet)
        @adjustment_hash = {}
        @jumbo_adjustment = {}
        @cash_out = {}
        @program_ids = []
        fixed_key = ''
        ltv_key = ''
        @sheet = sheet
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
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end
              if (term.nil? && @title.include?("ARM"))
                term = 0
              end

              # interest type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              # conforming
              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              # freddie_mac
              freddie_mac = false
              if @title.include?("Freddie Mac")
                freddie_mac = true
              end

              # fannie_mae
              fannie_mae = false
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                fannie_mae = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, sheet_name: sheet)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        # Adjustments
        (38..62).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(39)
          @ltv_arm_data = sheet_data.row(54)
          if row.compact.count >= 1
            (0..18).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
              if value.present?
                if value == "LTV Based Adjustments for 20/25/30 Yr Fixed Jumbo Products"
                  @adjustment_hash["LoanType/Term/FICO/LTV"] = {}
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"] = {}
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["20"] = {}
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["25"] = {}
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["30"] = {}

                  @cash_out["RefinanceOption/LTV"] = {}
                end
                if value == "LTV Based Adjustments for 15 Yr Fixed and All ARM Jumbo Products"
                  @jumbo_adjustment["LoanType/Term/FICO/LTV"] = {}
                  @jumbo_adjustment["LoanType/Term/FICO/LTV"]["Fixed"] = {}
                  @jumbo_adjustment["LoanType/Term/FICO/LTV"]["Fixed"]["15"] = {}
                  @jumbo_adjustment["LoanType/FICO/LTV"] = {}
                  @jumbo_adjustment["LoanType/FICO/LTV"]["ARM"] = {}

                  @jumbo_adjustment["LoanType/RefinanceOption/Term/FICO/LTV"] = {}
                  @jumbo_adjustment["LoanType/RefinanceOption/Term/FICO/LTV"]["Fixed"] = {}

                  @jumbo_adjustment["LoanType/RefinanceOption/FICO/LTV"] = {}
                  @jumbo_adjustment["LoanType/RefinanceOption/FICO/LTV"]["ARM"] = {}
                end
                # LTV Based Adjustments for 20/25/30 Yr Fixed Jumbo Products
                if (r >= 40 && r <= 45 && cc == 3)
                  if value.include?(">")
                    ltv_key = value.tr('>= ','')+"-Inf"
                  else
                    ltv_key = get_value value
                  end
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["20"][ltv_key] = {}
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["25"][ltv_key] = {}
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["30"][ltv_key] = {}
                end
                if r >= 40 && r <= 45 && cc > 3 && cc <= 14
                  if @ltv_data[cc-2].include?("<")
                    fixed_key = "0-"+ @ltv_data[cc-2].tr('<>= ','')
                  else
                    fixed_key = get_value @ltv_data[cc-2]
                  end
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["20"][ltv_key][fixed_key] = {}
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["25"][ltv_key][fixed_key] = {}
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["30"][ltv_key][fixed_key] = {}
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["20"][ltv_key][fixed_key] = value
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["25"][ltv_key][fixed_key] = value
                  @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["30"][ltv_key][fixed_key] = value
                end
                if r == 46 && cc == 2
                  @cash_out["RefinanceOption/LTV"]["Purchase"] = {}
                end
                if r == 46 && cc >= 4 && cc <= 14
                  if @ltv_data[cc-2].include?("<")
                    fixed_key = "0-"+ @ltv_data[cc-2].tr('<>= ','')
                  else
                    fixed_key = get_value @ltv_data[cc-2]
                  end
                  @cash_out["RefinanceOption/LTV"]["Purchase"][fixed_key] = {}
                  @cash_out["RefinanceOption/LTV"]["Purchase"][fixed_key] = value
                end
                if r == 47 && cc == 2
                  @cash_out["RefinanceOption/LTV"]["Cash Out"] = {}
                end
                if r == 47 && cc >= 4 && cc <= 14
                  if @ltv_data[cc-2].include?("<")
                    fixed_key = "0-"+ @ltv_data[cc-2].tr('<>= ','')
                  else
                    fixed_key = get_value @ltv_data[cc-2]
                  end
                  @cash_out["RefinanceOption/LTV"]["Cash Out"][fixed_key] = {}
                  @cash_out["RefinanceOption/LTV"]["Cash Out"][fixed_key] = value
                end
                if r == 48 && cc == 2
                  @cash_out["RefinanceOption/LTV"]["Rate and Term"] = {}
                end
                if r == 48 && cc >= 4 && cc <= 14
                  if @ltv_data[cc-2].include?("<")
                    fixed_key = "0-"+ @ltv_data[cc-2].tr('<>= ','')
                  else
                    fixed_key = get_value @ltv_data[cc-2]
                  end
                  @cash_out["RefinanceOption/LTV"]["Rate and Term"][fixed_key] = {}
                  @cash_out["RefinanceOption/LTV"]["Rate and Term"][fixed_key] = value
                end
                if r == 50 && cc == 2
                  @cash_out["RefinanceOption/LTV"]["Non-Owner Occupied"] = {}
                end
                if r == 50 && cc >= 4 && cc <= 14
                  if @ltv_data[cc-2].include?("<")
                    fixed_key = "0-"+ @ltv_data[cc-2].tr('<>= ','')
                  else
                    fixed_key = get_value @ltv_data[cc-2]
                  end
                  @cash_out["RefinanceOption/LTV"]["Non-Owner Occupied"][fixed_key] = {}
                  @cash_out["RefinanceOption/LTV"]["Non-Owner Occupied"][fixed_key] = value
                end
                # LTV Based Adjustments for 15 Yr Fixed and All ARM Jumbo Products
                if r >= 55 && r <= 60 && cc == 3
                  if value.include?(">")
                    ltv_key = value.tr('>= ','')+"-Inf"
                  else
                    ltv_key = get_value value
                  end
                  @jumbo_adjustment["LoanType/Term/FICO/LTV"]["Fixed"]["15"][ltv_key] = {}
                  @jumbo_adjustment["LoanType/FICO/LTV"]["ARM"][ltv_key] = {}
                end
                if r >= 55 && r <= 60 && cc > 3 && cc <= 14
                  if @ltv_arm_data[cc-2].include?("<")
                    fixed_key = "0-"+ @ltv_arm_data[cc-2].tr('<>= ','')
                  else
                    fixed_key = get_value @ltv_arm_data[cc-2]
                  end
                  @jumbo_adjustment["LoanType/Term/FICO/LTV"]["Fixed"]["15"][ltv_key][fixed_key] = {}
                  @jumbo_adjustment["LoanType/Term/FICO/LTV"]["Fixed"]["15"][ltv_key][fixed_key] = value
                  @jumbo_adjustment["LoanType/FICO/LTV"]["ARM"][ltv_key][fixed_key] = {}
                  @jumbo_adjustment["LoanType/FICO/LTV"]["ARM"][ltv_key][fixed_key] = value
                end
                if r == 61 && cc == 2
                  @jumbo_adjustment["LoanType/RefinanceOption/Term/FICO/LTV"]["Fixed"]["Purchase"] = {}
                  @jumbo_adjustment["LoanType/RefinanceOption/FICO/LTV"]["ARM"]["Purchase"] = {}
                  @jumbo_adjustment["LoanType/RefinanceOption/Term/FICO/LTV"]["Fixed"]["Purchase"]["15"] = {}
                end
                if r == 61 && cc > 3 && cc <= 14
                  if @ltv_arm_data[cc-2].include?("<")
                    fixed_key = "0-"+ @ltv_arm_data[cc-2].tr('<>= ','')
                  else
                    fixed_key = get_value @ltv_arm_data[cc-2]
                  end
                  @jumbo_adjustment["LoanType/RefinanceOption/Term/FICO/LTV"]["Fixed"]["Purchase"]["15"][fixed_key] = {}
                  @jumbo_adjustment["LoanType/RefinanceOption/FICO/LTV"]["ARM"]["Purchase"][fixed_key] = {}
                  @jumbo_adjustment["LoanType/RefinanceOption/Term/FICO/LTV"]["Fixed"]["Purchase"]["15"][fixed_key] = value
                  @jumbo_adjustment["LoanType/RefinanceOption/FICO/LTV"]["ARM"]["Purchase"][fixed_key] = value
                end
                if r == 62 && cc == 2
                  @jumbo_adjustment["LoanType/RefinanceOption/Term/FICO/LTV"]["Fixed"]["Cash Out"] = {}
                  @jumbo_adjustment["LoanType/RefinanceOption/FICO/LTV"]["ARM"]["Cash Out"] = {}
                  @jumbo_adjustment["LoanType/RefinanceOption/Term/FICO/LTV"]["Fixed"]["Cash Out"]["15"] = {}
                end
                if r == 62 && cc > 3 && cc <= 14
                  if @ltv_arm_data[cc-2].include?("<")
                    fixed_key = "0-"+ @ltv_arm_data[cc-2].tr('<>= ','')
                  else
                    fixed_key = get_value @ltv_arm_data[cc-2]
                  end
                  @jumbo_adjustment["LoanType/RefinanceOption/Term/FICO/LTV"]["Fixed"]["Cash Out"]["15"][fixed_key] = {}
                  @jumbo_adjustment["LoanType/RefinanceOption/FICO/LTV"]["ARM"]["Cash Out"][fixed_key] = {}
                  @jumbo_adjustment["LoanType/RefinanceOption/Term/FICO/LTV"]["Fixed"]["Cash Out"]["15"][fixed_key] = value
                  @jumbo_adjustment["LoanType/RefinanceOption/FICO/LTV"]["ARM"]["Cash Out"][fixed_key] = value
                end
              end
              rescue
                raise "value is nil at row = #{r} and column = #{cc}"
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@jumbo_adjustment,@cash_out]
        create_adjust(adjustment,sheet)
        create_program_association_with_adjustment(@sheet)
      end
    end
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def high_balance_extra
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @programs_ids = []
    xlsx.sheets.each do |sheet|
      if (sheet == "High Balance Extra")
        sheet_data = xlsx.sheet(sheet)
        @program_ids = []
        @adjustment_hash = {}
        @sub_hash = {}
        @cash_out = {}
        @bal_data = []
        @sub_data = []
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        key = ''
        bal_data = ''
        sub_data = ''
        @sheet = sheet
        (1..23).each do |r|
          row = sheet_data.row(r)
          if (row.compact.include?("High Balance Extra 30 Yr Fixed"))
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 2 + max_column*6 # (3 / 9 / 15)
              # title
              @title = sheet_data.cell(r,cc)

              # term
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end


              # rate type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end
              # High Balance
              jumbo_high_balance = false
              if @title.include?("High Balance")
                jumbo_high_balance = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
              # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase", arm_basic: arm_basic, sheet_name: sheet, jumbo_high_balance: jumbo_high_balance)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = r + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        # Adjustments
        (25..44).each do |r|
          row = sheet_data.row(r)
          @bal_data = sheet_data.row(27)
          @sub_data = sheet_data.row(41)
          if row.compact.count >= 1
            (0..9).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Pricing Adjustments"
                  @adjustment_hash["LoanSize/FICO/LTV"] = {}
                  @adjustment_hash["LoanSize/FICO/LTV"]["High Balance"] = {}
                end
                if value == "Cashout (adjustments are cumulative)"
                  @cash_out["RefinanceOption/FICO/LTV"] = {}
                  @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
                end
                if value == "Sub Financing (adjustments are cumulative)"
                  @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                  @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                end
                # All High Balance Extra Loans
                if r >= 28 && r <= 32 && cc == 2
                  if value.include?(">")
                    ltv_key = value.tr('>=','')+"-Inf"
                  else
                    ltv_key = get_value value
                  end
                  @adjustment_hash["LoanSize/FICO/LTV"]["High Balance"][ltv_key] = {}
                end
                if r >= 28 && r <= 32 && cc > 3 && cc <= 9
                  if @bal_data[cc-2].include?("<")
                    bal_data = "0-"+ @bal_data[cc-2].tr('<= ','')
                  else
                    bal_data = get_value @bal_data[cc-2]
                  end
                  @adjustment_hash["LoanSize/FICO/LTV"]["High Balance"][ltv_key][bal_data] = {}
                  @adjustment_hash["LoanSize/FICO/LTV"]["High Balance"][ltv_key][bal_data] = value
                end
                # Cashout Adjustments
                if r >= 34 && r <= 38 && cc == 2
                  if value.include?(">")
                    ltv_key = value.tr('>=','')+"-Inf"
                  else
                    ltv_key = get_value value
                  end
                  @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][ltv_key] = {}
                end
                if r >= 34 && r <= 38 && cc > 3 && cc <= 9
                  if @bal_data[cc-2].include?("<")
                    bal_data = "0-"+ @bal_data[cc-2].tr('<= ','')
                  else
                    bal_data = get_value @bal_data[cc-2]
                  end
                  @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][ltv_key][bal_data] = {}
                  @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][ltv_key][bal_data] = value
                end

                # Subordinate Financing Adjustments
                if r >= 42 && r <= 44 && cc == 2
                  if value.include?("<")
                    ltv_key = "0-"+ value.tr('<= ','')
                  else
                    ltv_key = get_value value
                  end
                  @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key] = {}
                end
                if r >= 42 && r <= 44 && cc == 3
                  cltv_key = get_value value
                  @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key] = {}
                end
                if r >= 42 && r <= 44 && cc > 3 && cc <= 5
                  if @sub_data[cc-2].include?(">")
                    sub_data = @sub_data[cc-2].tr('>= ','')+"-Inf"
                  else
                    sub_data = get_value @sub_data[cc-2]
                  end
                  @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key][sub_data] = {}
                  @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key][sub_data] = value
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@sub_hash,@cash_out]
        create_adjust(adjustment,sheet)
        create_program_association_with_adjustment(@sheet)
      end
    end
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def freddie_arms
    @program_ids = []
    @allAdjustments = {}
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Freddie ARMs")
        @sheet = sheet
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
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # rate type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              if @title.include?("3-1 ARM") || @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              # conforming
              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              # freddie_mac
              freddie_mac = false
              if @title.include?("Freddie Mac")
                freddie_mac = true
              end

              # fannie_mae
              fannie_mae = false
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                fannie_mae = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
              # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, sheet_name: sheet)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
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
              @title = "LoanSize/LoanType/FICO/LTV"

              unless @block_hash.has_key?(@title)
                @block_hash[@title] = {}
                @block_hash[@title]["Conforming"] = {}
                @block_hash[@title]["Conforming"]["ARM"] = {}
              end

              @block_hash["PropertyType"] = {} unless @block_hash.has_key?("PropertyType")
              @block_hash["MiscAdjuster"] = {} unless @block_hash.has_key?("MiscAdjuster")
              @block_hash["LockDay"] = {} unless @block_hash.has_key?("LockDay")
              @block_hash["LTV"] = {} unless @block_hash.has_key?("LTV")

              unless @block_hash.has_key?("LoanPurpose/RefinanceOption/LTV")
                @block_hash["LoanPurpose/RefinanceOption/LTV"] = {}
                @block_hash["LoanPurpose/RefinanceOption/LTV"]["Purchase"] = {}
                @block_hash["LoanPurpose/RefinanceOption/LTV"]["Purchase"]["Rate & Term"] = {}
                @block_hash["RefinanceOption/LTV"] = {}
                @block_hash["RefinanceOption/LTV"]["Cash Out"] = {}
              end

              key = ''
              another_key = ''
              keyOfHash = ''
              # for Misc Adjusters
              first_key   = ''
              second_key  = ''
              third_key   = ''
              final_key   = ''
              f1_key      = ''
              f2_key      = ''
              loan_amount = ''

              (0..47).each do |max_row|
                @data = []
                (7..19).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = index
                  value = sheet_data.cell(rrr,ccc)
                  # implementation of first key
                  if rrr.eql?(63)
                    # for 2nd table
                    @title = sheet_data.cell(rrr,cc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Cash Out"] = {}
                    end
                  elsif rrr.eql?(69)
                    # for 3rd table
                    previous_title = @title = sheet_data.cell(rrr,ccc - 4) unless previous_title == @title
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      first_key  = "LPMI/RefinanceOption/LTV"
                      second_key = "LPMI/PropertyType/LTV"
                      final_key  = "LPMI/FICO/LTV"
                      @block_hash[@title][first_key] = {}
                      @block_hash[@title][first_key][true] = {}
                      @block_hash[@title][second_key] = {}
                      @block_hash[@title][second_key][true] = {}
                      @block_hash[@title][final_key] = {}
                      @block_hash[@title][final_key][true] = {}
                    end
                  elsif rrr.eql?(81) && index == 7
                    # for Number Of Units
                    @title = sheet_data.cell(rrr,(ccc - 4))
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                    end
                  elsif rrr.eql?(81) && index == 13
                    @another_title = sheet_data.cell(rrr+1,index)
                    unless @block_hash.has_key?(@another_title)
                      @block_hash[@another_title] = {}
                    end
                  elsif rrr.eql?(84) && index.eql?(7)
                    # for Subordinate Financing
                    @title = sheet_data.cell(rrr,ccc - 4)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Subordinate Financing"] = {}
                    end
                  elsif rrr.eql?(89) && index.eql?(7)
                    # for Misc Adjusters
                    @title = sheet_data.cell(rrr,ccc - 4)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(94) && index.eql?(7)
                    # for Non Owner Occupied
                    @title = sheet_data.cell(rrr,ccc - 4)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                    @block_hash[@title]["Non-Owner Occupied"] = {}
                  end

                  # implementation of second key inside first key
                  if rrr > 52 && rrr < 60 && index == 7 && value
                    # for 1st table
                    key = get_value(value)
                    @block_hash[@title]["Conforming"]["ARM"][key] = {} unless @block_hash[@title]["Conforming"]["ARM"].has_key?(key)
                  elsif rrr > 62 && rrr < 67 && index == 7 && value
                    # for 2nd table
                    key = get_value(value)
                    @block_hash[@title]["Cash Out"][key] = {} unless @block_hash[@title]["Cash Out"].has_key?(key)
                  elsif (69..79).to_a.include?(rrr) && index == 7 && value
                    if(68..74).to_a.include?(rrr) && index == 7 && value
                      key = sheet_data.cell(rrr,ccc - 2)
                      if ["Rate & Term Refi", "Cash Out"].include?(key)
                        key = key.eql?("Rate & Term Refi") ? "Rate & Term" : "Cash Out"
                        @block_hash[@title][first_key][true][key] = {}
                      else
                        @block_hash[@title][second_key][true][key] = {}
                      end
                    elsif(76..79).to_a.include?(rrr) && index == 7 && value
                      key = sheet_data.cell(rrr,ccc - 2)
                      key = key.eql?("<= 85") ? set_range(key) : get_value(key)
                      @block_hash[@title][final_key][true][key] = {}
                    else
                      @block_hash[@title][key] = {}
                    end
                  else
                    if [82,83].include?(rrr) && index == 7
                      # for Number Of Units
                      key = sheet_data.cell(rrr,ccc - 1)
                      @block_hash[@title][key] = {}
                    elsif (rrr >= 82 && rrr <= 89) && index.eql?(15)
                      # for Loan Size Adjustments
                      loan_amount = get_value(sheet_data.cell(rrr, ccc))
                      @block_hash[@another_title][loan_amount] = {}
                    end

                    if (85..88).to_a.include?(rrr) && ccc < 13
                      # for Subordinate Financing
                      if index.eql?(7)
                        key = sheet_data.cell(rrr,ccc - 1)
                        key = set_range(key).nil? ? get_value(key) : set_range(key)
                        @block_hash[@title]["Subordinate Financing"][key] = {} unless @block_hash[@title]["Subordinate Financing"].has_key?(key)
                      elsif index.eql?(8)
                        keyOfHash = sheet_data.cell(rrr,ccc - 1)
                        keyOfHash = set_range(keyOfHash).nil? ? get_value(keyOfHash) : set_range(keyOfHash)
                        @block_hash[@title]["Subordinate Financing"][key][keyOfHash] = {}
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
                          first_key  = set_range(key.split(" LTV")[0])
                        end
                      end
                    end
                  end

                  # implementation of third key inside second key with value
                  if rrr > 52 && rrr < 60 && index > 7 && value
                    diff_of_row = rrr - 52
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 760") ? set_range(hash_key) : get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Conforming"]["ARM"][key][hash_key] = value unless @block_hash[@title]["Conforming"]["ARM"][key].has_key?(hash_key)
                    end
                  elsif rrr > 62 && rrr < 67 && index > 7 && value
                    # for 2nd table
                    hash_key = sheet_data.cell(rrr - (max_row + 1),ccc)
                    hash_key = hash_key.eql?(">= 760") ? set_range(hash_key) : get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Cash Out"][key][hash_key] = value unless @block_hash[@title]["Cash Out"][key].has_key?(hash_key)
                    end
                  elsif rrr >= 69 && index >= 7 && value
                    if (rrr <= 79)
                      # for 3rd table (69..79).to_a
                      diff_of_row = rrr - 68
                      hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                      hash_key = set_range(hash_key).nil? ? get_value(hash_key) : set_range(hash_key)
                      if (69..74).to_a.include?(rrr)
                        if(69..74).to_a.include?(rrr) && index >= 7 && value
                          if [69,74].include?(rrr)
                            @block_hash[@title][first_key][true][key][hash_key] = value
                          else
                            @block_hash[@title][second_key][true][key][hash_key] = value
                          end
                        end
                      elsif(76..79).to_a.include?(rrr) && index >= 7 && value
                        @block_hash[@title][final_key][true][key][hash_key] = value
                      end
                    elsif [82,83].include?(rrr) && [9,10,11].include?(index)
                      # for Number Of Units
                      diff_of_row = rrr - 81
                      hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                      hash_key = (hash_key.eql?("<=80") || hash_key.eql?("> 85")) ? set_range(hash_key) : get_value(hash_key)
                      @block_hash[@title][key][hash_key] = value if hash_key.present?
                    elsif (82..89).to_a.include?(rrr) && ccc > 15 && value
                      #for Loan Size Adjustments
                      diff_of_row  = rrr - 81
                      loan_purpose = sheet_data.cell(rrr-diff_of_row,ccc)
                      @block_hash[@another_title][loan_amount][loan_purpose] = value
                    end

                    if (85..88).to_a.include?(rrr) && [10,11].include?(ccc)
                      # for Subordinate Financing
                      diff_of_row = rrr - 84
                      hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                      hash_key = hash_key.eql?(">= 720") ? set_range(hash_key) : get_value(hash_key)
                      @block_hash[@title]["Subordinate Financing"][key][keyOfHash][hash_key] = value if hash_key.present?
                    end

                    if (89..93).to_a.include?(rrr) && ccc.eql?(11)
                      # for Misc Adjusters
                      if rrr.eql?(89)
                        @block_hash[@title][first_key][second_key][third_key] = value
                      elsif rrr.eql?(90)
                        @block_hash["LTV"][first_key] = value
                      elsif [91, 92].include?(rrr)
                        misc_adjuster = sheet_data.cell(rrr,(ccc -5))
                        @block_hash["MiscAdjuster"][misc_adjuster] = value
                      elsif rrr.eql?(93)
                        @block_hash["PropertyType"]["Manufactured Home"] = value
                      end
                    end

                    if [94,95,96].include?(rrr) && ccc.eql?(11)
                      #for Non Owner Occupied
                      occupancy_key = sheet_data.cell(rrr,(ccc -4))
                      occupancy_key = set_range(occupancy_key).nil? ? get_value(occupancy_key) : set_range(occupancy_key)
                      @block_hash[@title]["Non-Owner Occupied"][occupancy_key] = value if key && value
                    end

                    if (91..94).to_a.include?(rrr)
                      # for Super Conforming
                      if index.eql?(19)
                        has_key  = sheet_data.cell(rrr,ccc - 3)
                        has_key  = has_key.split(" ")[2] + has_key.split(" ")[3] unless [91, 92].include?(rrr)
                        hash_key = set_range(has_key.split(" ")[4]).nil? ? get_value(has_key.split(" ")[4]) : set_range(has_key.split(" ")[4]) if [91, 92].include?(rrr)
                        hash_key = set_range(has_key).nil? ? get_value(has_key) : set_range(has_key) unless [91, 92].include?(rrr)
                        @block_hash["LoanPurpose/RefinanceOption/LTV"]["Purchase"]["Rate & Term"][hash_key] = value if [91, 92].include?(rrr)
                        @block_hash["RefinanceOption/LTV"]["Cash Out"][hash_key] = value unless [91, 92].include?(rrr)
                      end
                    end

                    if rrr.eql?(96) && index.eql?(19)
                      @block_hash["LockDay"]["30"] = value
                      @block_hash["LockDay"]["45"] = value
                      @block_hash["LockDay"]["60"] = value
                    end
                  end
                end

                @allAdjustments[@title] = @block_hash[@title]
                if @another_title
                  @allAdjustments[@another_title] = @block_hash[@another_title]
                  @allAdjustments["PropertyType"] = @block_hash["PropertyType"]
                  @allAdjustments["MiscAdjuster"] = @block_hash["MiscAdjuster"]
                  @allAdjustments["PropertyType/LTV"] = @block_hash["PropertyType/LTV"]
                  @allAdjustments["LockDay"] = @block_hash["LockDay"]
                  @allAdjustments["LTV"] = @block_hash["LTV"]
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
    make_adjust(@allAdjustments, @sheet)
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def conforming_arms
    @program_ids = []
    @allAdjustments = {}
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Conforming ARMs")
        @sheet = sheet
        sheet_data = xlsx.sheet(sheet)

        (1..47).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || row.compact.include?("Fannie Mae 10-1 ARM (5-2-5) High Balance")
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)

              @title = sheet_data.cell(r,cc)
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              if @title.include?("3-1 ARM") || @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              freddie_mac = false
              if @title.include?("Freddie Mac")
                freddie_mac = true
              end

              conforming = false
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                conforming = true
              end

              fannie_mae = false
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                fannie_mae = true
              end

              # High Balance
              jumbo_high_balance = false
              if @title.include?("High Balance")
                jumbo_high_balance = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @program_ids << @program.id
              # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, jumbo_high_balance: jumbo_high_balance, sheet_name: sheet, arm_basic: arm_basic)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        # adjustments
        previous_title = nil
        @another_title = nil
        modified_keys  = get_table_keys
        data = get_table_keys
        (49..95).each do |r|
          row = sheet_data.row(r)
          # r == 52 / 68 / 81 / 84 / 89 / 94
          rr = r #+ 1 # (r == 53) / (r == 69) / (r == 82) / (r == 90) / (r == 95)
          max_column_section = row.compact.count - 1
          (0..max_column_section).each do |max_column|
            cc = 3 + max_column * 9 # (2 / 11)
            @title = sheet_data.cell(r,cc)
            @block_hash = {}

            if(@title.eql?("All Conforming ARMs (Does not include DU Refi Plus)"))
              @title = "LoanSize/LoanType/FICO/LTV"
              @block_hash[@title] = {}
              @block_hash[@title]["Conforming"] = {}
              @block_hash[@title]["Conforming"]["ARM"] = {}
              @block_hash["PropertyType"] = {} unless @block_hash.has_key?("PropertyType")
              @block_hash["MiscAdjuster"] = {} unless @block_hash.has_key?("MiscAdjuster")
              @block_hash["LockDay"] = {} unless @block_hash.has_key?("LockDay")
              #  for high balance key
              @block_hash["LoanSize/LoanPurpose/RefinanceOption/LTV"] = {} unless @block_hash.has_key?("LoanSize/LoanPurpose/RefinanceOption/LTV")
              @block_hash["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High Balance"] = {} unless @block_hash["LoanSize/LoanPurpose/RefinanceOption/LTV"].has_key?("High Balance")
              @block_hash["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High Balance"]["Purchase"] = {} unless @block_hash["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High Balance"].has_key?("Purchase")
              @block_hash["LoanSize/RefinanceOption/LTV"] = {} unless @block_hash.has_key?("LoanSize/RefinanceOption/LTV")
              @block_hash["LoanSize/RefinanceOption/LTV"]["High Balance"] = {} unless @block_hash["LoanSize/RefinanceOption/LTV"].has_key?("High Balance")

              key = ''
              another_key = ''
              keyOfHash = ''
              # for Misc Adjusters
              first_key   = ''
              second_key  = ''
              third_key   = ''
              final_key   = ''
              f1_key      = ''
              f2_key      = ''
              loan_amount = ''

              (0..47).each do |max_row|
                @data = []
                (7..19).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = index
                  value = sheet_data.cell(rrr,ccc)
                  # implementation of first key
                  if rrr.eql?(63)
                    # for 2nd table
                    @title = sheet_data.cell(rrr,cc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Cash Out"] = {}
                    end
                  elsif rrr.eql?(68)
                    # for 3rd table
                    previous_title = @title = sheet_data.cell(rrr,ccc - 4) unless previous_title == @title
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      first_key  = "LPMI/RefinanceOption/LTV"
                      second_key = "LPMI/PropertyType/LTV"
                      final_key  = "LPMI/FICO/LTV"
                      @block_hash[@title][first_key] = {}
                      @block_hash[@title][first_key][true] = {}
                      @block_hash[@title][second_key] = {}
                      @block_hash[@title][second_key][true] = {}
                      @block_hash[@title][final_key] = {}
                      @block_hash[@title][final_key][true] = {}
                    end
                  elsif rrr.eql?(81) && index == 13
                    # for Loan Size Adjustments
                    @another_title = sheet_data.cell(rrr,index)
                    unless @block_hash.has_key?(@another_title)
                      @block_hash[@another_title] = {}
                    end
                  elsif rrr.eql?(81) && index == 7
                    # for Subordinate Financing
                    @title = sheet_data.cell(rrr,ccc - 4)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Subordinate Financing"] = {}
                    end
                  elsif rrr.eql?(87) && index.eql?(7)
                    # for Non Owner Occupied
                    @title = sheet_data.cell(rrr,ccc - 4)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                    @block_hash[@title]["Non-Owner Occupied"] = {}
                  elsif rrr.eql?(91) && index.eql?(7)
                    # for Misc Adjusters
                    @title = "PropertyType/LTV"
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  end

                  # implementation of second key inside first key
                  if rrr > 52 && rrr < 61 && index == 7 && value
                    # for 1st table
                    key = get_value(value)
                    @block_hash[@title]["Conforming"]["ARM"][key] = {} unless @block_hash[@title]["Conforming"]["ARM"].has_key?(key)
                  elsif rrr > 62 && rrr < 66 && index == 7 && value
                    # for 2nd table
                    key = get_value(value)
                    @block_hash[@title]["Cash Out"][key] = {} unless @block_hash[@title]["Cash Out"].has_key?(key)
                  elsif (68..79).to_a.include?(rrr) && index == 7 && value
                    if(68..72).to_a.include?(rrr) && index == 7 && value
                      key = sheet_data.cell(rrr,ccc - 2)
                      if "Rate & Term Refi".eql?(key)
                        key = "Rate & Term"
                        @block_hash[@title][first_key][true][key] = {}
                      else
                        @block_hash[@title][second_key][true][key] = {}
                      end
                    elsif(75..78).to_a.include?(rrr) && index == 7 && value
                      key = sheet_data.cell(rrr,ccc - 2)
                      key = key.eql?("<= 85") ? set_range(key) : get_value(key)
                      @block_hash[@title][final_key][true][key] = {}
                    end
                  else
                    if (81..84).to_a.include?(rrr) && ccc < 12
                      # for Subordinate Financing
                      if index.eql?(7)
                        key = sheet_data.cell(rrr,ccc - 2)
                        key = key.eql?("<=65") ? set_range(key) : get_value(key)
                        @block_hash[@title]["Subordinate Financing"][key] = {} unless @block_hash[@title]["Subordinate Financing"].has_key?(key)
                      elsif index.eql?(8)
                        keyOfHash = sheet_data.cell(rrr,ccc - 2)
                        keyOfHash = get_value(keyOfHash)
                        @block_hash[@title]["Subordinate Financing"][key][keyOfHash] = {}
                      end
                    end

                    if (rrr >= 81 && rrr <= 88) && index.eql?(15)
                      # for Loan Size Adjustments
                      loan_amount = get_value(sheet_data.cell(rrr, ccc))
                      @block_hash[@another_title][loan_amount] = {}
                    end

                    if [87,88,89].include?(rrr) && [7].include?(ccc)
                      #for Non Owner Occupied
                      diff_of_column = ccc - 6
                      hash_key = sheet_data.cell(rrr,(ccc -diff_of_column))
                      hash_key = hash_key.eql?("> 80") ? set_range(hash_key) : get_value(hash_key)
                      key = hash_key
                      @block_hash[@title]["Non-Owner Occupied"][hash_key] = {} if hash_key.present?
                    end

                    if [89,91].include?(rrr)
                      # for High Balance
                      if rrr.eql?(89) && index.eql?(16)
                        @block_hash["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High Balance"]["Purchase"]["Rate and Term"] = {} unless @block_hash["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High Balance"]["Purchase"].has_key?("Rate and Term")
                      elsif rrr.eql?(91) && index.eql?(16)
                        @block_hash["LoanSize/RefinanceOption/LTV"]["High Balance"]["Cash Out"] = {} unless @block_hash["LoanSize/RefinanceOption/LTV"]["High Balance"].has_key?("Cash Out")
                      end
                    end

                    if (91..95).to_a.include?(rrr)
                      # for Misc Adjusters
                      if index.eql?(7)
                        key = sheet_data.cell(rrr,ccc - 2)
                        if key && key.eql?("Attached Condo > 75 LTV")
                          first_key = key.split(" >")[0].split(" ").last
                          @block_hash[@title][first_key] = {}
                          second_key = key.split(" ")[3] + ".01"
                        elsif key && key.eql?(">90 LTV")
                          first_key  = key.split(" ")[1]
                          @block_hash[@title][first_key] = {}
                          second_key = key.split(">")[1].split(" ").first
                        end
                      end
                    end
                  end

                  # implementation of third key inside second key with value
                  if rrr > 52 && rrr < 61 && index > 7 && value
                    diff_of_row = rrr - 52
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Conforming"]["ARM"][key][hash_key] = value unless @block_hash[@title]["Conforming"]["ARM"][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 62 && rrr < 67 && index > 7 && value
                    # for 2nd table
                    hash_key = sheet_data.cell(rrr - (max_row + 1),ccc)
                    hash_key = get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Cash Out"][key][hash_key] = value unless @block_hash[@title]["Cash Out"][key].has_key?(hash_key)
                    end
                  elsif rrr >= 68 && index >= 7 && value
                    if(rrr <= 78)
                      diff_of_row = rrr - 67
                      hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                      hash_key = set_range(hash_key).nil? ? get_value(hash_key) : set_range(hash_key)
                      if (68..72).to_a.include?(rrr)
                        if(68..72).to_a.include?(rrr) && index >= 7 && value
                          if rrr.eql?(68)
                            @block_hash[@title][first_key][true][key][hash_key] = value
                          else
                            @block_hash[@title][second_key][true][key][hash_key] = value
                          end
                        end
                      elsif(75..78).to_a.include?(rrr) && index >= 7 && value
                        @block_hash[@title][final_key][true][key][hash_key] = value
                      end
                    elsif (81..88).to_a.include?(rrr) && ccc > 15 && value
                      #for Loan Size Adjustments
                      diff_of_row  = rrr - 80
                      loan_purpose = sheet_data.cell(rrr-diff_of_row,ccc)
                      @block_hash[@another_title][loan_amount][loan_purpose] = value
                    end

                    if (81..84).to_a.include?(rrr) && [9,10].include?(ccc)
                      # for Subordinate Financing
                      diff_of_row = rrr - 80
                      hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                      hash_key = hash_key.eql?(">= 720") ? set_range(hash_key) : get_value(hash_key)
                      @block_hash[@title]["Subordinate Financing"][key][keyOfHash][hash_key] = value if hash_key.present?
                    end

                    if [87,88,89].include?(rrr) && [9].include?(ccc)
                      @block_hash[@title]["Non-Owner Occupied"][key] = value if key && value
                    end

                    if (89..92).to_a.include?(rrr)
                      # for High Balance
                      ltv_value = set_range(sheet_data.cell(rrr,ccc - 1).split("LTV ")[-1]) if index.eql?(19)
                      if [89,90].include?(rrr) && index.eql?(19)
                        @block_hash["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High Balance"]["Purchase"]["Rate and Term"][ltv_value] = value
                      elsif [91,92].include?(rrr) && index.eql?(19)
                        @block_hash["LoanSize/RefinanceOption/LTV"]["High Balance"]["Cash Out"][ltv_value] = value
                      end
                    end

                    if (91..96).to_a.include?(rrr) && ccc.eql?(9)
                      # for Misc Adjusters
                      if rrr.eql?(93)
                        @block_hash[@title][first_key][second_key] = value
                      elsif rrr.eql?(96)
                        ltv_value = sheet_data.cell(rrr, ccc - 4).split(">")[-1].split("LTV")[0] + " - #{Float::INFINITY}"
                        @block_hash[@title][first_key][ltv_value] = value
                      elsif [91,92].include?(rrr)
                        property_type_value = sheet_data.cell(rrr, ccc - 4)
                        @block_hash["PropertyType"][property_type_value] = value
                      elsif [94,95].include?(rrr)
                        misc_adjuster_value = sheet_data.cell(rrr, ccc - 4)
                        @block_hash["MiscAdjuster"][misc_adjuster_value] = value
                      end
                    end

                    if rrr.eql?(94) && ccc.eql?(17)
                      @block_hash["LockDay"]["30"] = value
                      @block_hash["LockDay"]["45"] = value
                      @block_hash["LockDay"]["60"] = value
                    end
                  end
                end

                @allAdjustments[@title] = @block_hash[@title]
                if @another_title
                  @allAdjustments[@another_title] = @block_hash[@another_title]
                  @allAdjustments["LoanSize/LoanPurpose/RefinanceOption/LTV"] = @block_hash["LoanSize/LoanPurpose/RefinanceOption/LTV"]
                  @allAdjustments["LoanSize/RefinanceOption/LTV"] = @block_hash["LoanSize/RefinanceOption/LTV"]
                  @allAdjustments["PropertyType/LTV"] = @block_hash["PropertyType/LTV"]
                  @allAdjustments["PropertyType"] = @block_hash["PropertyType"]
                  @allAdjustments["MiscAdjuster"] = @block_hash["MiscAdjuster"]
                  @allAdjustments["LockDay"] = @block_hash["LockDay"]
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
    make_adjust(@allAdjustments, @sheet)
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def homeready
    program_ids = []
    @allAdjustments = {}
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "HomeReady")
        @sheet = sheet
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


              # term
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # rate type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              conforming = false
              fannie_mae = false
              if @title.include?("Fannie Mae")
                conforming = true
                fannie_mae = true
              end
              fannie_mae_home_ready = false
              if @title.include?("Fannie Mae HomeReady")
                fannie_mae_home_ready = true
              end

              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              program_ids << @program.id
               # Loan Limit Type
              if @title.include?("Non-Conforming")
                @program.loan_limit_type << "Non-Conforming"
              end
              if @title.include?("Conforming")
                @program.loan_limit_type << "Conforming"
              end
              if @title.include?("Jumbo")
                @program.loan_limit_type << "Jumbo"
              end
              if @title.include?("High Balance")
                @program.loan_limit_type << "High Balance"
              end
              @program.save
              @program.update(term: term,loan_type: loan_type, arm_basic: arm_basic, loan_purpose: "Purchase", fannie_mae: fannie_mae, fannie_mae_home_ready: fannie_mae_home_ready, conforming: conforming, sheet_name: sheet)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        # adjustments
        previous_title = nil
        @another_title = nil
        modified_keys  = get_table_keys
        data = get_table_keys
        (81..128).each do |r|
          row    = sheet_data.row(r)
          # r == 52 / 68 / 81 / 84 / 89 / 94
          rr = r #+ 1 # (r == 53) / (r == 69) / (r == 82) / (r == 90) / (r == 95)
          max_column_section = row.compact.count - 1
          (0..max_column_section).each do |max_column|
            cc = 3 + max_column * 9 # (2 / 11)
            @title = sheet_data.cell(r,cc)
            @block_hash = {}
            if(@title.eql?("All Fixed Conforming (does not apply to terms ≤ 15yrs)"))
              @title = "LoanSize/LoanType/Term/FICO/LTV"
              @block_hash[@title] = {}
              @block_hash[@title]["Conforming"] = {}
              @block_hash[@title]["Conforming"]["Fixed"] = {}
              @block_hash[@title]["Conforming"]["Fixed"]["0-15"] = {}
              @block_hash["PropertyType"] = {} unless @block_hash.has_key?("PropertyType")
              @block_hash["LockDay"] = {} unless @block_hash.has_key?("LockDay")

              key = ''
              another_key = ''
              keyOfHash = ''
              # for Misc Adjusters
              first_key   = ''
              second_key  = ''
              third_key   = ''
              final_key   = ''
              f1_key      = ''
              f2_key      = ''
              loan_amount = ''

              (0..48).each do |max_row|
                @data = []
                (3..19).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = index
                  value = sheet_data.cell(rrr,ccc)
                  # implementation of first key
                  if rrr.eql?(91)
                    # for Cash-Out
                    @title = sheet_data.cell(rrr,cc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Cash Out"] = {}
                    end
                  elsif rrr.eql?(98) && index == 3
                    # for Lender Paid MI Adjustments
                    previous_title = @title = sheet_data.cell(rrr,ccc) unless previous_title == @title
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      first_key  = "LPMI/RefinanceOption/LTV"
                      second_key = "LPMI/PropertyType/LTV"
                      final_key  = "LPMI/Term/FICO/LTV"
                      @block_hash[@title][first_key] = {}
                      @block_hash[@title][first_key][true] = {}
                      @block_hash[@title][second_key] = {}
                      @block_hash[@title][second_key][true] = {}
                      @block_hash[@title][final_key] = {}
                      @block_hash[@title][final_key][true] = {}
                    end
                  elsif rrr.eql?(113) && index == 3
                    # for Subordinate Financing
                    @title = sheet_data.cell(rrr,ccc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Subordinate Financing"] = {}
                    end
                  elsif rrr.eql?(113) && index == 13
                    # for Loan Size Adjustments
                    @another_title = sheet_data.cell(rrr,index)
                    unless @block_hash.has_key?(@another_title)
                      @block_hash[@another_title] = {}
                    end
                  elsif rrr.eql?(119) && index == 3
                    # for Misc Adjusters
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(122) && index.eql?(3)
                    # for Non Owner Occupied
                    @another_title = sheet_data.cell(rrr,ccc)
                    @block_hash[@another_title] = {} unless @block_hash.has_key?(@another_title)
                    @block_hash[@another_title]["Non Owner Occupied"] = {}
                  elsif rrr.eql?(127) && index.eql?(13)
                    # for Adjustment Caps
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  end

                  #implementation of second key inside first key
                  if rrr > 80 && rrr < 89 && index == 7 && value
                    key = get_value(value)
                    @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key] = {} unless @block_hash[@title]["Conforming"]["Fixed"]["0-15"].has_key?(key)
                  elsif rrr > 90 && rrr < 94 && index == 7 && value
                    # for 2nd table
                    key = get_value(value)
                    @block_hash[@title]["Cash Out"][key] = {} unless @block_hash[@title]["Cash Out"].has_key?(key)
                  elsif (rrr > 97) && (rrr < 111)
                    # for Lender Paid MI Adjustments
                    if index == 5 && value
                      if ["Rate & Term Refi", "Cash Out"].include?(value)
                        f1_key = "Rate & Term"
                        @block_hash[@title][first_key][true][f1_key] = {}
                      elsif ["Manufactured Home", "2nd Home", "3-4 Unit", "Non Owner Occupied"].include?(value)
                        f2_key = value
                        @block_hash[@title][second_key][true][f2_key] = {}
                      else
                        @block_hash[@title][final_key][true]["0 - 20"] = {} if value.eql?("≤ 20 Yr Term")
                        @block_hash[@title][final_key][true]["20 - Infinity"] = {} if value.eql?("> 20 Yr Term")
                        key = "0 - 20" if value.eql?("≤ 20 Yr Term")
                        key = "20 - Infinity" if value.eql?("> 20 Yr Term")
                      end
                    elsif index == 6 && rrr < 154 && value
                      another_key = value.eql?("≤ 85") ? set_range(value) : value
                      @block_hash[@title][final_key][true][key][another_key] = {} if another_key
                    end
                  end

                  if (113..118).to_a.include?(rrr) && ccc < 12
                    # for Subordinate Financing
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      key = get_value(key)
                      @block_hash[@title]["Subordinate Financing"][key] = {} unless @block_hash[@title].has_key?(key)
                    elsif index.eql?(7)
                      keyOfHash = sheet_data.cell(rrr,ccc)
                      keyOfHash = get_value(keyOfHash)
                      @block_hash[@title]["Subordinate Financing"][key][keyOfHash] = {}
                    end
                  end

                  if (rrr >= 114 && rrr <= 121) && index.eql?(15)
                    # for Loan Size Adjustments
                    loan_amount = get_value(sheet_data.cell(rrr, ccc))
                    @block_hash[@another_title][loan_amount] = {}
                  end

                  if (119..121).to_a.include?(rrr)
                    # for Misc Adjusters
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      if key && key.eql?("Attached Condo > 75 LTV (>15yr Term)")
                        first_key = key.split(" >")[0].split(" ").last
                        @block_hash[@title][first_key] = {}
                        second_key = key.split(" ")[3] + ".01"
                        @block_hash[@title][first_key][second_key] = {}
                        third_key = key.split(" ")[5].split("(>")[1].split("yr")[0] + ".01"
                      elsif key && key.eql?(">90 LTV")
                        first_key  = key.split(" ")[1]
                        @block_hash[@title][first_key] = {}
                        second_key = key.split(">")[1].split(" ").first
                      end
                    end
                  end

                  if [122,123,124].include?(rrr) && [7].include?(ccc)
                    #for Non Owner Occupied
                    hash_key = sheet_data.cell(rrr,ccc)
                    hash_key = key = (hash_key.eql?("> 80") ? set_range(hash_key) : get_value(hash_key))
                    @block_hash[@another_title]["Non Owner Occupied"][hash_key] = {} if hash_key.present?
                  end

                  if (127..129).to_a.include?(rrr) && @title
                    # for Adjustment Caps
                    if index.eql?(17)
                      another_key = sheet_data.cell(rrr,ccc)
                      @block_hash[@title][another_key] = {} if another_key
                    end
                  end

                  # implementation of third key inside second key with value
                  if rrr > 80 && rrr < 89 && index > 7 && value
                    diff_of_row = rrr - 80
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 760") ? set_range(hash_key) : get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key][hash_key] = value unless @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 90 && rrr < 94 && index > 7 && value
                    # for 2nd table
                    diff_of_row = rrr - 80
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 760") ? set_range(hash_key) : get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Cash Out"][key][hash_key] = value unless @block_hash[@title]["Cash Out"][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 97 && rrr <= 110 && index >= 7 && value
                    # for Lender Paid MI Adjustments
                    diff_of_row = rrr - 97
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = set_range(hash_key).nil? ? get_value(hash_key) : set_range(hash_key)
                    if (98..100).to_a.include?(rrr)
                      if [98].include?(rrr)
                        @block_hash[@title][first_key][true][f1_key][hash_key] = value
                      else
                        @block_hash[@title][second_key][true][f2_key][hash_key] = value
                      end
                    else
                      if rrr > 101
                        @block_hash[@title][final_key][true][key][another_key][hash_key] = value
                      end
                    end
                  end

                  if (113..118).to_a.include?(rrr) && ccc > 9 && ccc < 12 && value
                    # for Subordinate Financing
                    diff_of_row = rrr - 112
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 720") ? set_range(hash_key) : get_value(hash_key)
                    @block_hash[@title]["Subordinate Financing"][key][keyOfHash][hash_key] = value if hash_key.present?
                  end

                  if (114..121).to_a.include?(rrr) && ccc > 15 && value
                    #for Loan Size Adjustments
                    diff_of_row  = rrr - 113
                    loan_purpose = sheet_data.cell(rrr-diff_of_row,ccc)
                    @block_hash[@another_title][loan_amount][loan_purpose] = value
                  end

                  if (119..121).to_a.include?(rrr) && ccc == 11
                    #for Misc Adjusters
                    if rrr.eql?(120)
                      @block_hash[@title][first_key][second_key][third_key] = value
                    else
                      first_key = sheet_data.cell(rrr,ccc - 5)
                      @block_hash["PropertyType"][first_key] = value
                    end
                  end

                  if [122,123,124].include?(rrr) && [11].include?(ccc)
                    #for Non Owner Occupied
                    @block_hash[@another_title]["Non Owner Occupied"][key] = value if key && value
                  end

                  if (127..129).to_a.include?(rrr)
                    # for Adjustment Caps
                    if (18..19).to_a.include?(ccc)
                      diff_of_row = rrr - 126
                      has_key = sheet_data.cell((rrr-diff_of_row),ccc)
                      unless @block_hash[@title][another_key].has_key?(has_key)
                        @block_hash[@title][another_key][has_key] = value if another_key.present?
                      else
                        has_key = has_key + "1"
                        @block_hash[@title][another_key][has_key] = value if another_key.present?
                      end
                    end
                  end

                  if (rrr.eql?(122) || rrr.eql?(123)) && ccc.eql?(19)
                    @block_hash["MiscAdjuster"] = value if rrr.eql?(122)
                    @block_hash["PropertyType"] = value if rrr.eql?(123)
                  end

                  if rrr.eql?(126) && ccc.eql?(8)
                    @block_hash["LockDay"]["30"] = value
                    @block_hash["LockDay"]["45"] = value
                    @block_hash["LockDay"]["60"] = value
                  end
                end

                @allAdjustments[@title] = @block_hash[@title]
                if @another_title
                  @allAdjustments[@another_title] = @block_hash[@another_title]
                  @allAdjustments["PropertyType"] = @block_hash["PropertyType"]
                  @allAdjustments["LockDay"] = @block_hash["LockDay"]
                  @allAdjustments["MiscAdjuster"] = @block_hash["MiscAdjuster"]
                  @allAdjustments["PropertyType"] = @block_hash["PropertyType"]
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
    make_adjust(@allAdjustments, @sheet)
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def homeready_hb
    program_ids = []
    @allAdjustments = {}
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "HomeReady HB")
        @sheet = sheet
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
               # term
              term = nil
              program_heading = @title.split
              if @title.include?("10yr") || @title.include?("10 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("15yr") || @title.include?("15 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("20yr") || @title.include?("20 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("25yr") || @title.include?("25 Yr")
                term = @title.scan(/\d+/)[0]
              elsif @title.include?("30yr") || @title.include?("30 Yr")
                term = @title.scan(/\d+/)[0]
              end

              # rate type
              if @title.include?("Fixed")
                loan_type = "Fixed"
              elsif @title.include?("ARM")
                loan_type = "ARM"
              elsif @title.include?("Floating")
                loan_type = "Floating"
              elsif @title.include?("Variable")
                loan_type = "Variable"
              else
                loan_type = nil
              end

              # rate arm
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                arm_basic = @title.scan(/\d+/)[0].to_i
              end

              conforming = false
              fannie_mae = false
              if @title.include?("Fannie Mae")
                conforming = true
                fannie_mae = true
              end
              fannie_mae_home_ready = false
              if @title.include?("Fannie Mae HomeReady")
                fannie_mae_home_ready = true
              end
              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              program_ids << @program.id
              @program.update(term: term,loan_type: loan_type, arm_basic: arm_basic, loan_purpose: "Purchase", fannie_mae: fannie_mae, fannie_mae_home_ready: fannie_mae_home_ready, conforming: conforming, sheet_name: sheet)
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[main_key][key] = {}
                  else
                    if @program.lock_period.length <= 3
                      @program.lock_period << 15*c_i
                      @program.save
                    end
                    @block_hash[main_key][key][15*c_i] = value
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        # adjustments
        previous_title = nil
        @another_title = nil
        modified_keys  = get_table_keys
        data = get_table_keys
        (80..127).each do |r|
          row    = sheet_data.row(r)
          # r == 52 / 68 / 81 / 84 / 89 / 94
          rr = r #+ 1 # (r == 53) / (r == 69) / (r == 82) / (r == 90) / (r == 95)
          max_column_section = row.compact.count - 1
          (0..max_column_section).each do |max_column|
            cc = 3 + max_column * 9 # (2 / 11)
            @title = sheet_data.cell(r,cc)
            @block_hash = {}

            if(@title.eql?("All Fixed Conforming\n(does not apply to terms ≤ 15yrs)"))
              @title = "LoanSize/LoanType/Term/FICO/LTV"
              @block_hash[@title] = {}
              @block_hash[@title]["Conforming"] = {}
              @block_hash[@title]["Conforming"]["Fixed"] = {}
              @block_hash[@title]["Conforming"]["Fixed"]["0-15"] = {}
              @block_hash["PropertyType"] = {} unless @block_hash.has_key?("PropertyType")
              @block_hash["LockDay"] = {} unless @block_hash.has_key?("LockDay")
              @block_hash["MiscAdjuster"] = {} unless @block_hash.has_key?("MiscAdjuster")
              @block_hash["LoanSize/LoanType/LTV"] = {} unless @block_hash.has_key?("LoanSize/LoanType/LTV")
              @block_hash["LoanSize/LoanType"] = {} unless @block_hash.has_key?("LoanSize/LoanType")

              key = ''
              another_key = ''
              keyOfHash = ''
              # for Misc Adjusters
              first_key   = ''
              second_key  = ''
              third_key   = ''
              final_key   = ''
              f1_key      = ''
              f2_key      = ''
              loan_amount = ''

              (0..47).each do |max_row|
                @data = []
                (3..19).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = index
                  value = sheet_data.cell(rrr,ccc)
                  # implementation of first key
                  if rrr.eql?(90)
                    # for Cash-Out
                    @title = sheet_data.cell(rrr,cc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Cash Out"] = {}
                    end
                  elsif rrr.eql?(97) && index == 3
                    # for Lender Paid MI Adjustments
                    previous_title = @title = sheet_data.cell(rrr,ccc) unless previous_title == @title
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      first_key  = "LPMI/RefinanceOption/LTV"
                      second_key = "LPMI/PropertyType/LTV"
                      final_key  = "LPMI/Term/FICO/LTV"
                      @block_hash[@title][first_key] = {}
                      @block_hash[@title][first_key][true] = {}
                      @block_hash[@title][second_key] = {}
                      @block_hash[@title][second_key][true] = {}
                      @block_hash[@title][final_key] = {}
                      @block_hash[@title][final_key][true] = {}
                    end
                  elsif rrr.eql?(112) && index == 3
                    # for Subordinate Financing
                    @title = sheet_data.cell(rrr,ccc)
                    unless @block_hash.has_key?(@title)
                      @block_hash[@title] = {}
                      @block_hash[@title]["Subordinate Financing"] = {}
                    end
                  elsif rrr.eql?(112) && index == 13
                    # for Loan Size Adjustments
                    @another_title = sheet_data.cell(rrr,index)
                    unless @block_hash.has_key?(@another_title)
                      @block_hash[@another_title] = {}
                    end
                  elsif rrr.eql?(118) && index == 3
                    # for Misc Adjusters
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  elsif rrr.eql?(126) && index.eql?(13)
                    # for Adjustment Caps
                    @title = sheet_data.cell(rrr,ccc)
                    @block_hash[@title] = {} unless @block_hash.has_key?(@title)
                  end

                  #implementation of second key inside first key
                  if rrr > 79 && rrr < 88 && index == 7 && value
                    key = get_value(value)
                    @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key] = {} unless @block_hash[@title]["Conforming"]["Fixed"]["0-15"].has_key?(key)
                  elsif rrr > 89 && rrr < 93 && index == 7 && value
                    # for 2nd table
                    key = get_value(value)
                    @block_hash[@title]["Cash Out"][key] = {} unless @block_hash[@title]["Cash Out"].has_key?(key)
                  elsif (rrr > 96) && (rrr < 110)
                    # for Lender Paid MI Adjustments
                    if index == 5 && value
                      if ["Rate & Term Refi", "Cash Out"].include?(value)
                        f1_key = "Rate & Term"
                        @block_hash[@title][first_key][true][f1_key] = {}
                      elsif ["Manufactured Home", "2nd Home", "3-4 Unit", "Non Owner Occupied"].include?(value)
                        f2_key = value
                        @block_hash[@title][second_key][true][f2_key] = {}
                      else
                        @block_hash[@title][final_key][true]["0 - 20"] = {} if value.eql?("≤ 20 Yr Term")
                        @block_hash[@title][final_key][true]["20 - Infinity"] = {} if value.eql?("> 20 Yr Term")
                        key = "0 - 20" if value.eql?("≤ 20 Yr Term")
                        key = "20 - Infinity" if value.eql?("> 20 Yr Term")
                      end
                    elsif index == 6 && rrr < 110 && value
                      another_key = value.eql?("≤ 85") ? set_range(value) : value
                      @block_hash[@title][final_key][true][key][another_key] = {} if another_key
                    end
                  end

                  if (112..117).to_a.include?(rrr) && ccc < 12
                    # for Subordinate Financing
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      key = get_value(key)
                      @block_hash[@title]["Subordinate Financing"][key] = {} unless @block_hash[@title].has_key?(key)
                    elsif index.eql?(7)
                      keyOfHash = sheet_data.cell(rrr,ccc)
                      keyOfHash = get_value(keyOfHash)
                      @block_hash[@title]["Subordinate Financing"][key][keyOfHash] = {}
                    end
                  end

                  if (rrr >= 113 && rrr <= 120) && index.eql?(15)
                    # for Loan Size Adjustments
                    loan_amount = get_value(sheet_data.cell(rrr, ccc))
                    @block_hash[@another_title][loan_amount] = {}
                  end

                  if (118..123).to_a.include?(rrr)
                    # for Misc Adjusters
                    if index.eql?(6)
                      key = sheet_data.cell(rrr,ccc)
                      if key && key.eql?("Attached Condo > 75 LTV (>15yr Term)")
                        first_key = key.split(" >")[0].split(" ").last
                        @block_hash[@title][first_key] = {}
                        second_key = key.split(" ")[3] + ".01"
                        @block_hash[@title][first_key][second_key] = {}
                        third_key = key.split(" ")[5].split("(>")[1].split("yr")[0] + ".01"
                      elsif key && key.eql?(">90 LTV")
                        first_key  = key.split(" ")[1]
                        @block_hash[@title][first_key] = {}
                        second_key = key.split(">")[1].split(" ").first
                      end
                    end
                  end

                  if (126..128).to_a.include?(rrr) && @title
                    # for Adjustment Caps
                    if index.eql?(17)
                      another_key = sheet_data.cell(rrr,ccc)
                      @block_hash[@title][another_key] = {} if another_key
                    end
                  end

                  # implementation of third key inside second key with value
                  if rrr > 79 && rrr < 88 && index > 7 && value
                    diff_of_row = rrr - 79
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 760") ? set_range(hash_key) : get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key][hash_key] = value unless @block_hash[@title]["Conforming"]["Fixed"]["0-15"][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 89 && rrr < 93 && index > 7 && value
                    # for 2nd table
                    diff_of_row = rrr - 79
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 760") ? set_range(hash_key) : get_value(hash_key)
                    if hash_key.present?
                      @block_hash[@title]["Cash Out"][key][hash_key] = value unless @block_hash[@title]["Cash Out"][key].has_key?(hash_key)
                    end
                  end

                  if rrr > 96 && rrr <= 109 && index >= 7 && value
                    # for Lender Paid MI Adjustments
                    diff_of_row = rrr - 96
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = set_range(hash_key).nil? ? get_value(hash_key) : set_range(hash_key)
                    if (97..99).to_a.include?(rrr)
                      if rrr.eql?(97)
                        @block_hash[@title][first_key][true][f1_key][hash_key] = value
                      elsif [98, 99].include?(rrr)
                        @block_hash[@title][second_key][true][f2_key][hash_key] = value
                      end
                    else
                      if rrr > 100
                        @block_hash[@title][final_key][true][key][another_key][hash_key] = value
                      end
                    end
                  end

                  if (112..117).to_a.include?(rrr) && ccc > 9 && ccc < 12 && value
                    # for Subordinate Financing
                    diff_of_row = rrr - 111
                    hash_key = sheet_data.cell((rrr - diff_of_row),ccc)
                    hash_key = hash_key.eql?("≥ 720") ? set_range(hash_key) : get_value(hash_key)
                    @block_hash[@title]["Subordinate Financing"][key][keyOfHash][hash_key] = value if hash_key.present?
                  end

                  if (113..120).to_a.include?(rrr) && ccc > 15 && value
                    #for Loan Size Adjustments
                    diff_of_row  = rrr - 112
                    loan_purpose = sheet_data.cell(rrr-diff_of_row,ccc)
                    @block_hash[@another_title][loan_amount][loan_purpose] = value
                  end

                  if (118..123).to_a.include?(rrr) && ccc == 11
                    #for Misc Adjusters
                    if rrr.eql?(119)
                      @block_hash[@title][first_key][second_key][third_key] = value
                    else
                      first_key = sheet_data.cell(rrr,ccc - 5)
                      if ["High Balance Loan Adjustment - Fixed", "High Balance Loan Adjustment - ARM LTV ≤ 75", "High Balance Loan Adjustment - ARM LTV > 75"].include?(first_key)
                        if "High Balance Loan Adjustment - Fixed".include?(first_key)
                          @block_hash["LoanSize/LoanType"]["High Balance"] = {}
                          @block_hash["LoanSize/LoanType"]["High Balance"]["Fixed"] = {}
                          @block_hash["LoanSize/LoanType"]["High Balance"]["Fixed"] = value
                        elsif "High Balance Loan Adjustment - ARM LTV ≤ 75".include?(first_key)
                          @block_hash["LoanSize/LoanType/LTV"]["High Balance"] = {}
                          @block_hash["LoanSize/LoanType/LTV"]["High Balance"]["ARM"] = {}
                          @block_hash["LoanSize/LoanType/LTV"]["High Balance"]["ARM"]["0-75"] = value
                        elsif "High Balance Loan Adjustment - ARM LTV > 75".include?(first_key)
                          @block_hash["LoanSize/LoanType/LTV"]["High Balance"]["ARM"]["75-Infinity"] = value
                        end
                      else
                        @block_hash["PropertyType"][first_key] = value if ["2-4 Units", "Manufactured Home"].include?(first_key)
                      end
                    end
                  end

                  if (126..128).to_a.include?(rrr)
                    # for Adjustment Caps
                    if (18..19).to_a.include?(ccc)
                      diff_of_row = rrr - 125
                      has_key = sheet_data.cell((rrr-diff_of_row),ccc)
                      unless @block_hash[@title][another_key].has_key?(has_key)
                        @block_hash[@title][another_key][has_key] = value if another_key.present?
                      else
                        has_key = has_key + "1"
                        @block_hash[@title][another_key][has_key] = value if another_key.present?
                      end
                    end
                  end

                  if rrr.eql?(125) && ccc.eql?(8)
                    @block_hash["LockDay"]["30"] = value
                    @block_hash["LockDay"]["45"] = value
                    @block_hash["LockDay"]["60"] = value
                  end
                end

                @allAdjustments[@title] = @block_hash[@title]
                if @another_title
                  @allAdjustments[@another_title] = @block_hash[@another_title]
                  @allAdjustments["PropertyType"] = @block_hash["PropertyType"]
                  @allAdjustments["LockDay"] = @block_hash["LockDay"]
                  @allAdjustments["LoanSize/LoanType/LTV"] = @block_hash["LoanSize/LoanType/LTV"]
                  @allAdjustments["LoanSize/LoanType"] = @block_hash["LoanSize/LoanType"]
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
    make_adjust(@allAdjustments, @sheet)
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_import_file_path(@sheet_obj)
  end

  def programs
    @programs = @sheet_obj.programs
  end

  def single_program
  end

  private

  def get_sheet
    @sheet_obj = Sheet.find(params[:id])
  end

  def get_program
    @program = Program.find(params[:id])
  end

  def get_bank
    @bank = Bank.find(params[:id])
  end

  def get_titles
    return ["FICO/LTV Adjustments - Loan Amount ≤ $1MM", "State Adjustments", "FICO/LTV Adjustments - Loan Amount > $1MM", "Feature Adjustments", "Max Price"]
  end

  def all_lp
    data = Adjustment::ALL_IP

    return data
  end

  def high_bal_adjustment
    data = Adjustment::HIGH_BALANCE_ADJUSTMENT
    return data
  end

  def jumbo_series_i_adjustment
      data = Adjustment::JUMBO_SERIES_I_ADJUSTMENT
    return data
  end

  def dream_big_adjustment
    data = Adjustment::DREAM_BIG_ADJUSTMENT

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

  # def make_adjust(block_hash, p_ids)
  #   begin
  #     adjustment = Adjustment.create(data: block_hash)

  #     # assign for all projects
  #     p_ids.each do |id|
  #       program = Program.find(id)
  #       program.adjustments << adjustment
  #     end
  #   rescue Exception => e
  #     puts e
  #   end
  # end

  def make_adjust(block_hash, sheet)
    block_hash.keys.each do |key|
      unless ["Lender Paid MI Adj.", "Term/LTV/FICO"].include?(key)
        hash = {}
        hash[key] = block_hash[key]
        Adjustment.create(data: hash,sheet_name: sheet)
      else
        unless block_hash[key].empty?
          block_hash[key].keys.each do |s_key|
            h1 = {}
            h1[s_key] = block_hash[key][s_key]
            Adjustment.create(data: h1,sheet_name: sheet)
          end
        end
      end
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
    table_keys = Adjustment::MAIN_KEYS
    return table_keys
  end

  def get_value value1
    if value1.present?
      if (!value1.include?("$")) && ((value1.include?("≤")) || (value1.include?("<")))
        value1 = "0 - " + value1.split().last
      elsif (value1.include?("-")) && !value1.include?("$")
        # value1 = value1.split("-").first.squish
        value1 = value1
      elsif (value1.include?("≥"))
        value1 = value1.split("≥").last.squish
      elsif (value1.include?(">="))
        value1.split(">=").last.squish
      elsif (value1.include?(">"))
        value1.split(">").last.squish
      elsif (value1.include?("+"))
        value1.split("+").first
      elsif value1.include?("$") && !value1.include?("-")
        "0 - " + value1.split("$").last.gsub(/[\s,]/ ,"").squish
      elsif value1.include?("$") && value1.include?("-")
        if !value1.split(" - ").last.eql?("Conforming Limit")
          value1 = value1.split("$")[1].gsub(/[\s,]/ ,"") + value1.split("$")[-1].gsub(/[\s,]/ ,"")
        else
          value1 = value1.split(" - ").first.gsub("$", "").gsub(",", "") + " - " + value1.split(" - ").last.squish
        end
      else
        value1
      end
    end
  end

  def set_range value
    if value.split()[0].eql?("≤") || value.split()[0].eql?("<=") then
      value = "0 - " + value.split()[1]
    elsif [">","≥",">=", "+"].include?(value.split()[0]) then
      value.split()[1] + " - #{Float::INFINITY}"
    elsif [">","≥",">=", "+"].include?(value.split("")[-1])
      value.split("+")[0] + " - #{Float::INFINITY}"
    elsif value.include?(">")
      value.split(">")[-1] + " - #{Float::INFINITY}"
    elsif value.include?("<=")
      value = "0 - " + value.split("<=")[-1]
    end
  end

  def get_main_key heading
    heading.split(" ").each do |data|
      data.gsub!("#{data}", 'LoanType') if data.eql?("Fixed")
      data.gsub!("#{data}", 'Term') if data.eql?("terms")
    end
  end

  def create_program_association_with_adjustment(sheet)
    adjustment_list = Adjustment.where(sheet_name: sheet)
    program_list = Program.where(sheet_name: sheet)

    adjustment_list.each_with_index do |adj_ment, index|
      key_list = adj_ment.data.keys.first.split("/")
      program_filter1={}
      program_filter2={}
      include_in_input_values = false
      if key_list.present?
        key_list.each_with_index do |key_name, key_index|
          if (Program.column_names.include?(key_name.underscore))
            unless (Program.column_for_attribute(key_name.underscore).type.to_s == "boolean")
              program_filter1[key_name.underscore] = nil
            else
              if (Program.column_for_attribute(key_name.underscore).type.to_s == "boolean")
                program_filter2[key_name.underscore] = true
              end
            end
          else
            if(Adjustment::INPUT_VALUES.include?(key_name))
              include_in_input_values = true
            end
          end
        end

        if (include_in_input_values)
          program_list1 = program_list.where.not(program_filter1)
          program_list2 = program_list1.where(program_filter2)

          if program_list2.present?
            program_list2.map{ |program| program.adjustments << adj_ment unless program.adjustments.include?(adj_ment) }
          end
        end
      end
    end
  end
  def create_adjust(block_hash, sheet)
    block_hash.each do |hash|
      if hash.present?
        hash.each do |key|
          data = {}
          data[key[0]] = key[1]
          Adjustment.create(data: data,sheet_name: sheet)
        end
      end
    end
  end
end
