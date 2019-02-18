class ObAlliedMortgageGroupWholesale8570Controller < ApplicationController
  before_action :get_sheet, only: [:programs, :fha, :va, :conf_fixed]
  before_action :get_program, only: [:single_program]
  def index
    file = File.join(Rails.root,  'OB_Allied_Mortgage_Group_Wholesale8570.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "Cover")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "Allied Mortgage"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def fha
    file = File.join(Rails.root,  'OB_Allied_Mortgage_Group_Wholesale8570.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FHA")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @fha_adjustment = {}
        @loan_adj = {}
        first_key = ''
        #program
        (13..80).each do |r| 
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + (2+max_column) # 2, 7, 12, 17
              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != 3.125
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property @program
                @programs_ids << @program.id
                # @program.adjustments.destroy_all
                key = ''
                @block_hash = {}
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
                if @block_hash.keys.first.nil? || @block_hash.keys.first == "Rate"
                  @block_hash.shift
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end

        # Adjustments FHA
        (39..57).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (17..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "FHA & USDA Loan Level Adjustments"
                  @fha_adjustment["FHA/USDA/FICO"] = {}
                  @fha_adjustment["FHA/USDA/FICO"][true] = {}
                  @fha_adjustment["FHA/USDA/FICO"][true][true] = {}
                end
                if r >=41 && r <= 45 && cc == 17
                  first_key = get_value value
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @fha_adjustment["FHA/USDA/FICO"][true][true][first_key] = c_val
                end
                if r == 46 && cc == 17
                  @fha_adjustment["FHA/USDA/LoanPurpose"] = {}
                  @fha_adjustment["FHA/USDA/LoanPurpose"][true] = {}
                  @fha_adjustment["FHA/USDA/LoanPurpose"][true][true] = {}
                  @fha_adjustment["FHA/USDA/LoanPurpose"][true][true]["Refinance"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @fha_adjustment["FHA/USDA/LoanPurpose"][true][true]["Refinance"] = new_val
                end
                if r == 47 && cc == 17
                  @fha_adjustment["FHA/USDA"] = {}
                  @fha_adjustment["FHA/USDA"][true] = {}
                  @fha_adjustment["FHA/USDA"][true][true] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @fha_adjustment["FHA/USDA"][true][true] = new_val
                end
                if r == 48 && cc == 17
                  @fha_adjustment["FHA/USDA/LoanSize/FICO"] = {}
                  @fha_adjustment["FHA/USDA/LoanSize/FICO"][true] = {}
                  @fha_adjustment["FHA/USDA/LoanSize/FICO"][true][true] = {}
                  @fha_adjustment["FHA/USDA/LoanSize/FICO"][true][true]["High Balance"] = {}
                  @fha_adjustment["FHA/USDA/LoanSize/FICO"][true][true]["High Balance"]["0-680"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @fha_adjustment["FHA/USDA/LoanSize/FICO"][true][true]["High Balance"]["0-680"] = new_val
                end
                if r == 49 && cc == 17
                  @fha_adjustment["FHA/Streamline/CLTV"] = {}
                  @fha_adjustment["FHA/Streamline/CLTV"][true] = {}
                  @fha_adjustment["FHA/Streamline/CLTV"][true][true] = {}
                  @fha_adjustment["FHA/Streamline/CLTV"][true][true]["100-125"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @fha_adjustment["FHA/Streamline/CLTV"][true][true]["100-125"] = new_val
                end
                if r == 50 && cc == 17
                  @fha_adjustment["FHA/USDA/PropertyType"] = {}
                  @fha_adjustment["FHA/USDA/PropertyType"][true] = {}
                  @fha_adjustment["FHA/USDA/PropertyType"][true][true] = {}
                  @fha_adjustment["FHA/USDA/PropertyType"][true][true]["Gov'n Non Owner"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @fha_adjustment["FHA/USDA/PropertyType"][true][true]["Gov'n Non Owner"] = new_val
                end
                if r == 51 && cc == 17
                  @fha_adjustment["FHA/USDA/LoanAmount/FICO"] = {}
                  @fha_adjustment["FHA/USDA/LoanAmount/FICO"][true] = {}
                  @fha_adjustment["FHA/USDA/LoanAmount/FICO"][true][true] = {}
                  @fha_adjustment["FHA/USDA/LoanAmount/FICO"][true][true]["0-100k"] = {}
                  @fha_adjustment["FHA/USDA/LoanAmount/FICO"][true][true]["0-100k"]["0-640"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @fha_adjustment["FHA/USDA/LoanAmount/FICO"][true][true]["0-100k"]["0-640"] = new_val
                end
                if r == 52 && cc == 17
                  @fha_adjustment["FHA/Usda/State"] = {}
                  @fha_adjustment["FHA/Usda/State"][true] = {}
                  @fha_adjustment["FHA/Usda/State"][true][true] = {}
                  @fha_adjustment["FHA/Usda/State"][true][true]["NY"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @fha_adjustment["FHA/Usda/State"][true][true]["NY"] = new_val
                end
              end
            end
          end
        end

        # Adjustments INDICES
        (84..95).each do |r|
          row = sheet_data.row(r)
          @term_data = sheet_data.row(93)
          if (row.compact.count >= 1)
            (0..20).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Allied Wholesale Loan Amt Adj"
                  @loan_adj["LoanAmount"] = {}
                  @loan_adj["LoanType/Term"] = {}
                  @loan_adj["LoanType/Term"]["Fixed"] = {}
                  @loan_adj["LoanType/Term"]["ARM"] = {}
                end
                if r >= 87 && r <= 95 && cc == 17
                  if value.include?(">")
                    first_key = get_value value
                  else
                    first_key = value.sub('to','-').tr('$','')
                  end
                  @loan_adj["LoanAmount"][first_key] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @loan_adj["LoanAmount"][first_key] = new_val
                end
                if r == 94 && cc >= 2 && cc <= 6
                  first_key = @term_data[cc-2].tr('A-Za-z ','')
                  @loan_adj["LoanType/Term"]["Fixed"][first_key] = {}
                  @loan_adj["LoanType/Term"]["Fixed"][first_key] = value
                end
                if r == 94 && cc >= 7 && cc <= 9
                  first_key = @term_data[cc-2].tr('A-Za-z ','')
                  @loan_adj["LoanType/Term"]["ARM"][first_key] = {}
                  @loan_adj["LoanType/Term"]["ARM"][first_key] = value
                end
              end
            end
          end
        end
        adjustment = [@fha_adjustment,@loan_adj]
        make_adjust(adjustment,sheet)
        # create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_allied_mortgage_group_wholesale8570_path(@sheet_obj)
  end

  def va
    file = File.join(Rails.root,  'OB_Allied_Mortgage_Group_Wholesale8570.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "VA")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @loan_amount = {}
        primary_key = ''

        # programs
        (13..79).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 5))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|
              cc = (4*max_column) + (2+max_column)  # (2 / 7 / 12)
              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != 3.5 && @title != 3.125 && @title != "Loan Amount"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                # Program Property
                program_property @title
                @program.adjustments.destroy_all
                key = ''
                @block_hash = {}
                (1..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row +1
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      else
                        @block_hash[key][15*(c_i)] = value unless @block_hash[key].nil?
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
        end

        # VA loan level adjustment
        (15..91).each do |r|
          row = sheet_data.row(r)
          @term_data = sheet_data.row(89)
          if (row.compact.count >= 1)
            (0..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "VA Loan Level Adjustments"
                  @adjustment_hash["VA/FICO"] = {}
                  @adjustment_hash["VA/FICO"][true] = {}
                  @adjustment_hash["VA/LTV"] = {}
                  @adjustment_hash["VA/LTV"][true] = {}
                end
                if value == "Allied Wholesale Loan Amt Adj *"
                  @loan_amount["LoanAmount"] = {}
                  @loan_amount["LoanType/Term"] = {}
                  @loan_amount["LoanType/Term"]["Fixed"] = {}
                  @loan_amount["LoanType/Term"]["ARM"] = {}
                end
                # VA Loan Level Adjustments
                if r >=17 && r <= 21 && cc == 17
                  primary_key = get_value value
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @adjustment_hash["VA/FICO"][true][primary_key] = c_val
                end
                if r == 22 && cc == 17
                  @adjustment_hash["VA/LoanSize/FICO"] = {}
                  @adjustment_hash["VA/LoanSize/FICO"][true] = {}
                  @adjustment_hash["VA/LoanSize/FICO"][true]["High Balance"] = {}
                  @adjustment_hash["VA/LoanSize/FICO"][true]["High Balance"]["0-680"] = {}
                  cc == cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @adjustment_hash["VA/LoanSize/FICO"][true]["High Balance"]["0-680"] = new_val
                end
                if r == 23 && cc == 17
                  @adjustment_hash["VA/LTV"][true]["90-95"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @adjustment_hash["VA/LTV"][true]["90-95"] = new_val
                end
                if r == 24 && cc == 14
                  @adjustment_hash["VA/LTV"][true]["95-Inf"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @adjustment_hash["VA/LTV"][true]["95-Inf"] = new_val
                end
                if r == 25 && cc == 14
                  @adjustment_hash["State"] = {}
                  @adjustment_hash["State"]["NY"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @adjustment_hash["State"]["NY"] = new_val
                end
                if r >= 37 && r <= 45 && cc == 17
                  if value.include?("to")
                    primary_key = value.sub('to','-').tr('$><% ','')
                  else
                    primary_key = get_value value
                  end
                  @loan_amount["LoanAmount"][primary_key] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @loan_amount["LoanAmount"][primary_key] = new_val
                end
                if r == 90 && cc >= 2 && cc <= 6
                  primary_key = @term_data[cc-2].tr('A-Za-z ','')
                  @loan_amount["LoanType/Term"]["Fixed"][primary_key] = {}
                  @loan_amount["LoanType/Term"]["Fixed"][primary_key] = value
                end
                if r == 90 && cc >= 7 && cc <= 9
                  primary_key = @term_data[cc-2].tr('A-Za-z ','')
                  @loan_amount["LoanType/Term"]["ARM"][primary_key] = {}
                  @loan_amount["LoanType/Term"]["ARM"][primary_key] = value
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@loan_amount]
        make_adjust(adjustment,sheet)
        # create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_allied_mortgage_group_wholesale8570_path(@sheet_obj)
  end

  def conf_fixed
    file = File.join(Rails.root,  'OB_Allied_Mortgage_Group_Wholesale8570.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "CONF FIXED")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @cash_out = {}
        @subordinate_hash = {}
        @property_hash = {}
        @other_adjustment = {}
        @loan_amount = {}
        primary_key = ''
        secondary_key = ''
        ltv_key = ''
        cltv_key = ''

        #program
        (13..56).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|

              cc = 4*max_column + (2+max_column) # 2, 7, 12, 17

              @title = sheet_data.cell(r,cc)

              # term
              @term = nil
              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
                @term = 30
              elsif @title.include?("20 Year") || @title.include?("20 Yr")
                @term = 20
              elsif @title.include?("15 Year") || @title.include?("15 Yr")
                @term = 15
              elsif @title.include?("30/25 Yr")
                @term = 30
              elsif @title.include?("10 Yr")
                @term = 10
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

              # streamline
              if @title.include?("FHA")
                @streamline = true
                @fha = true
                @full_doc = true
              elsif @title.include?("VA")
                @streamline = true
                @va = true
                @full_doc = true
              elsif @title.include?("USDA")
                @streamline = true
                @usda = true
                @full_doc = true
              else
                @streamline = false
                @fha = false
                @va = false
                @usda = false
                @full_doc = false
              end

              # High Balance
              if @title.include?("High Bal")
                @jumbo_high_balance = true
              end

              # Program Category
              if @title.include?("C30/C25")
                @program_category = "C30/C25"
              elsif @title.include?("C20")
                @program_category = "C20"
              elsif @title.include?("C15")
                @program_category = "C15"
              elsif @title.include?("C30")
                @program_category = "C30"
              elsif @title.include?("C10")
                @program_category = "C10"
              elsif @title.include?("C30JLP")
                @program_category = "C30JLP"
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
              @program.update(term: @term,loan_type: loan_type,loan_purpose: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc)
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              # if @program.term.present?
              #   main_key = "Term/LoanType/InterestRate/LockPeriod"
              # else
              #   main_key = "InterestRate/LockPeriod"
              # end
              @block_hash = {}
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
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*c_i
                        @program.save
                      end
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              # if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
              #   @block_hash.values.first.shift
              # end
              if @block_hash.keys.first.nil? || @block_hash.keys.first == "Rate"
                @block_hash.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        # VA loan level adjustment
        (59..111).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(63)
          @sub_data = sheet_data.row(64)
          @term_data = sheet_data.row(110)
          if (row.compact.count >= 1)
            (0..20).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "LOAN LEVEL PRICE ADJUSTMENTS"
                  @adjustment_hash["FICO/LTV"] = {}
                  @cash_out["RefinanceOption/FICO/LTV"] = {}
                  @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
                  @property_hash["PropertyType/FICO/LTV"] = {}
                end
                if value == "SUBORDINATE FINANCING"
                  @subordinate_hash["FinancingType/LTV/CLTV/FICO"] = {}
                  @subordinate_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                  @other_adjustment["FICO/Term"] = {}
                  @loan_amount["LoanAmount"] = {}
                  @loan_amount["LoanType/Term"] = {}
                  @loan_amount["LoanType/Term"]["Fixed"] = {}
                  @loan_amount["LoanType/Term"]["ARM"] = {}
                end
                if r >=65 && r <= 71 && cc == 4
                  primary_key = get_value value
                  @adjustment_hash["FICO/LTV"][primary_key] = {}
                end
                if r >=65 && r <= 71 && cc >= 5 && cc <= 12
                  if @ltv_data[cc-2].include?("-")
                    secondary_key = @ltv_data[cc-2].tr('%','')
                  else
                    secondary_key = get_value @ltv_data[cc-2]
                  end
                  @adjustment_hash["FICO/LTV"][primary_key][secondary_key] = {}
                  @adjustment_hash["FICO/LTV"][primary_key][secondary_key] = value
                end
                # Subordinate Financing
                if r >= 65 && r <= 69 && cc == 13
                  if value.include?("to")
                    ltv_key = value.sub('to','-').tr('$><% ','')
                  elsif value.include?("Any") || value.include?("ANY")
                    ltv_key = value
                  else
                    ltv_key = get_value value
                  end
                  @subordinate_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key] = {}
                end
                if r >= 65 && r <= 69 && cc == 15 
                  if value.include?("to")
                    cltv_key = value.sub('to','-').tr('$><% ','')
                  elsif value.include?("CLTV")
                    cltv_key = value  
                  else
                    cltv_key = get_value value
                  end
                  @subordinate_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key] = {}
                end
                if r >= 65 && r <= 69 && cc >= 17 && cc <= 19
                  sub_key = get_value @sub_data[cc-2]
                  @subordinate_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key][sub_key] = {}
                  @subordinate_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key][sub_key] = value
                end
                if r == 70 && cc == 13
                  @other_adjustment["MiscAdjuster"] = {}
                  @other_adjustment["MiscAdjuster"][value] = {}
                  cc = cc + 7
                  new_val = sheet_data.cell(r,cc)
                  @other_adjustment["MiscAdjuster"][value] = new_val
                end
                if r >= 71 && r <= 72 && cc == 13
                  ltv_key = value.tr('A-Za-z)( ','')
                  @other_adjustment["FICO/Term"][ltv_key] = {}
                  @other_adjustment["FICO/Term"][ltv_key]["0-Inf"] = {}
                  cc = cc + 7
                  new_val = sheet_data.cell(r,cc)
                  @other_adjustment["FICO/Term"][ltv_key]["0-Inf"] = new_val
                end
                # Cashout
                if r >= 72 && r <= 78 && cc == 4
                  primary_key = get_value value
                  @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key] = {}
                end
                if r >= 72 && r <= 78 && cc >= 5 && cc <= 12
                  if @ltv_data[cc-2].include?("-")
                    secondary_key = @ltv_data[cc-2].tr('%','')
                  else
                    secondary_key = get_value @ltv_data[cc-2]
                  end
                  @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][secondary_key] = {}
                  @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][secondary_key] = value
                end
                if r == 73 && cc == 13
                  @other_adjustment["FreddieMac/FICO"] = {}
                  @other_adjustment["FreddieMac/FICO"][true] = {}
                  @other_adjustment["FreddieMac/FICO"][true]["640-679"] = {}
                  cc = cc + 7
                  new_val = sheet_data.cell(r,cc)
                  @other_adjustment["FreddieMac/FICO"][true]["640-679"] = new_val
                end
                if r == 74 && cc == 13
                  @other_adjustment["LoanSize/FICO"] = {}
                  @other_adjustment["LoanSize/FICO"]["High Balance"] = {}
                  @other_adjustment["LoanSize/FICO"]["High Balance"]["0-740"] = {}
                  cc = cc + 7
                  new_val = sheet_data.cell(r,cc)
                  @other_adjustment["LoanSize/FICO"]["High Balance"]["0-740"] = new_val
                end
                if r == 75 && cc == 13
                  @other_adjustment["LoanSize/RefinanceOption"] = {}
                  @other_adjustment["LoanSize/RefinanceOption"]["High Balance"] = {}
                  @other_adjustment["LoanSize/RefinanceOption"]["High Balance"]["Rate and Term"] = {}
                  cc = cc + 7
                  new_val = sheet_data.cell(r,cc)
                  @other_adjustment["LoanSize/RefinanceOption"]["High Balance"]["Rate and Term"] = new_val
                end
                if r == 76 && cc == 13
                  @other_adjustment["LoanSize/RefinanceOption"]["High Balance"]["Cash Out"] = {}
                  cc = cc + 7
                  new_val = sheet_data.cell(r,cc)
                  @other_adjustment["LoanSize/RefinanceOption"]["High Balance"]["Cash Out"] = new_val
                end
                if r == 77 && cc == 13
                  @other_adjustment["State"] = {}
                  @other_adjustment["State"]["NY"] = {}
                  cc = cc + 7
                  new_val = sheet_data.cell(r,cc)
                  @other_adjustment["State"]["NY"] = new_val
                end
                # PropertyType
                if r >= 79 && r <= 83 && cc == 4
                  if value == "Condo*"
                    primary_key = "Condo"
                  else
                    primary_key = value
                  end
                  @property_hash["PropertyType/FICO/LTV"][primary_key] = {}
                end
                if r >= 79 && r <= 83 && cc >= 5 && cc <= 12
                  if @ltv_data[cc-2].include?("-")
                    secondary_key = @ltv_data[cc-2].tr('%','')
                  else
                    secondary_key = get_value @ltv_data[cc-2]
                  end
                  @property_hash["PropertyType/FICO/LTV"][primary_key][secondary_key] = {}
                  @property_hash["PropertyType/FICO/LTV"][primary_key][secondary_key] = value
                end
                if r == 83 && cc == 13
                  @other_adjustment["FannieMaeProduct/FreddieMacProduct/FICO/LTV"] = {}
                  @other_adjustment["FannieMaeProduct/FreddieMacProduct/FICO/LTV"]["HomeReady"] = {}
                  @other_adjustment["FannieMaeProduct/FreddieMacProduct/FICO/LTV"]["HomeReady"]["HomePossible"] = {}
                  @other_adjustment["FannieMaeProduct/FreddieMacProduct/FICO/LTV"]["HomeReady"]["HomePossible"]["680-Inf"] = {}
                  @other_adjustment["FannieMaeProduct/FreddieMacProduct/FICO/LTV"]["HomeReady"]["HomePossible"]["680-Inf"]["80-Inf"] = {}
                  cc = cc + 7
                  new_val = sheet_data.cell(r,cc)
                  @other_adjustment["FannieMaeProduct/FreddieMacProduct/FICO/LTV"]["HomeReady"]["HomePossible"]["680-Inf"]["80-Inf"] = new_val
                end
                if r == 84 && cc == 13
                  @other_adjustment["FannieMaeProduct/FreddieMacProduct/FICO/LTV"]["HomeReady"]["HomePossible"]["0-680"] = {}
                  @other_adjustment["FannieMaeProduct/FreddieMacProduct/FICO/LTV"]["HomeReady"]["HomePossible"]["0-680"]["0-80"] = {}
                  cc = cc + 7
                  new_val = sheet_data.cell(r,cc)
                  @other_adjustment["FannieMaeProduct/FreddieMacProduct/FICO/LTV"]["HomeReady"]["HomePossible"]["0-680"]["0-80"] = new_val
                end
                if r >= 91 && r <= 99 && cc == 13
                  if value.include?("to")
                    ltv_key = value.sub('to','-').tr('$><% ','')
                  else
                    ltv_key = get_value value
                  end
                  @loan_amount["LoanAmount"][ltv_key] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @loan_amount["LoanAmount"][ltv_key] = new_val
                end
                if r == 111 && cc >= 7 && cc <= 11
                  first_key = @term_data[cc-2].tr('A-Za-z ','')
                  @loan_amount["LoanType/Term"]["Fixed"][first_key] = {}
                  @loan_amount["LoanType/Term"]["Fixed"][first_key] = value
                end
                if r == 111 && cc >= 12 && cc <= 14
                  first_key = @term_data[cc-2].tr('A-Za-z ','')
                  @loan_amount["LoanType/Term"]["ARM"][first_key] = {}
                  @loan_amount["LoanType/Term"]["ARM"][first_key] = value
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@cash_out,@subordinate_hash,@property_hash,@other_adjustment,@loan_amount]
        make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_allied_mortgage_group_wholesale8570_path(@sheet_obj)
  end

  def programs
    @programs = @sheet_obj.programs
  end

  def single_program
  end

  private

    def get_value value1
      if value1.present?
        if value1.include?("<=") || value1.include?("<")
          value1 = "0-"+value1.split("<=").last.tr('^0-9', '')
        elsif value1.include?(">") || value1.include?("+")
          value1 = value1.split(">").last.tr('^0-9', '')+"-Inf"
        else
          value1 = value1.tr('A-Z ','')
        end
      end
    end

    def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end

    def get_program
      @program = Program.find(params[:id])
    end

    def program_property value1
      # term
      if @program.program_name.include?("30 Year") || @program.program_name.include?("30Yr") || @program.program_name.include?("30 Yr") || @program.program_name.include?("30/25 Year")
        term = 30
      elsif @program.program_name.include?("20 Year")
        term = 20
      elsif @program.program_name.include?("15 Year")
        term = 15
      elsif @program.program_name.include?("10 Year")
        term = 10
      else
        term = nil
      end

      # Loan-Type
      if @program.program_name.include?("Fixed")
        loan_type = "Fixed"
      elsif @program.program_name.include?("ARM")
        loan_type = "ARM"
      elsif @program.program_name.include?("Floating")
        loan_type = "Floating"
      elsif @program.program_name.include?("Variable")
        loan_type = "Variable"
      else
        loan_type = nil
      end

      # Streamline Vha, Fha, Usda
      fha = false
      va = false
      usda = false
      streamline = false
      full_doc = false
      if @program.program_name.include?("FHA")
        streamline = true
        fha = true
        full_doc = true
      elsif @program.program_name.include?("VA")
        streamline = true
        va = true
        full_doc = true
      elsif @program.program_name.include?("USDA")
        streamline = true
        usda = true
        full_doc = true
      end

      # High Balance
      jumbo_high_balance = false
      if @program.program_name.include?("High Bal")
        jumbo_high_balance = true
      end

       # Program Category
      program_category = @program.program_name.split.last

      # Loan Limit Type
      if @program.program_name.include?("Non-Conforming")
        @program.loan_limit_type << "Non-Conforming"
      end
      if @program.program_name.include?("Conforming")
        @program.loan_limit_type << "Conforming"
      end
      if @program.program_name.include?("Jumbo")
        @program.loan_limit_type << "Jumbo"
      end
      if @program.program_name.include?("High Balance")
        @program.loan_limit_type << "High Balance"
      end
      @program.save
      @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline, jumbo_high_balance: jumbo_high_balance, program_category: program_category)
    end

    def make_adjust(block_hash, sheet)
      block_hash.each do |hash|
        hash.each do |key|
          data = {}
          data[key[0]] = key[1]
          Adjustment.create(data: data,sheet_name: sheet)
        end
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
end
