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
        @fha_adjustment = {}
        @indi_adjustment = {}
        @lock_adjustment = {}
        @avrg_adjustment = {}
        @alhs_adjustment = {}
        @programs_ids = []
        @key_data = []
        @key_data2 = []
        f_key = ''
        first_key = ''
        cc = ''
        ccc = ''
        c_val = ''

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
              end

              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
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
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
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
                  first_key = "FHA/USDA Loan"
                  @fha_adjustment[first_key] = {}
                end

                if r >=41 && r <= 52 && cc == 17
                  value = get_value value
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @fha_adjustment[first_key][value] = c_val
                end

                if r >=56 && r <= 57 && cc == 17
                  value = get_value value
                  ccc = cc + 2
                  c_val = sheet_data.cell(r,ccc)
                  @fha_adjustment[first_key][value] = c_val
                end

              end
            end
          end
        end

        # Adjustments INDICES
        (84..94).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (1..10).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "INDICES"
                  f_key = "INDICES"
                  @indi_adjustment[f_key] = {}
                end

                if r >=85 && r <= 89 && cc == 2
                  ccc = cc + 2
                  c_val = sheet_data.cell(r,ccc)
                  @indi_adjustment[f_key][value] = c_val
                end

                if value == "Lock Expiration Dates:"
                  first_key = "Lock/ExpirationDates:"
                  @lock_adjustment[first_key] = {}
                end

                if r >=85 && r <= 87 && cc == 7
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @lock_adjustment[first_key][value] = c_val
                end
              end
            end
          end
        end

        #Adjustments Avrage Price
        (92..94).each do |r|
          row = sheet_data.row(r)
          @key_data = sheet_data.row(93)
          @key_data2 = sheet_data.row(93)
          if (row.compact.count >= 1)
            (2..9).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Average Price Offer Rate (APOR) For Loans Locked During The Week Of 01/14/2019"
                  f_key = "AveragePrice"
                  @avrg_adjustment[f_key] = {}
                end

                if r == 93 && cc >= 2 && cc <= 9
                  rr = r+1
                  c_val = sheet_data.cell(rr,cc)
                  @avrg_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end


        #Adjustments Allied Holesale
        (84..95).each do |r|
          row = sheet_data.row(r)
          @key_data = sheet_data.row(93)
          @key_data2 = sheet_data.row(93)
          if (row.compact.count >= 1)
            (17..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Allied Wholesale Loan Amt Adj"
                  f_key = "AlliedWholesale/LoanAmount/adj"
                  @alhs_adjustment[f_key] = {}
                end

                if r >= 87 && r <= 95 && cc == 17
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @alhs_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end
        Adjustment.create(data: @fha_adjustment, sheet_name: sheet)
        Adjustment.create(data: @indi_adjustment, sheet_name: sheet)
        Adjustment.create(data: @lock_adjustment, sheet_name: sheet)
        Adjustment.create(data: @avrg_adjustment, sheet_name: sheet)
        Adjustment.create(data: @alhs_adjustment, sheet_name: sheet)
        create_program_association_with_adjustment(sheet)
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
        @valn_adjustment = {}
        @awlm_adjustment = {}
        @indi_adjustment = {}
        @avrg_adjustment = {}
        @lock_adjustment = {}
        primary_key = ''
        second_key = ''
        t = 1
        c_val = ''
        f_key = ''

        # programs
        (13..79).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|

              # max_column = max_column + 1

              cc = (4*max_column) + (2+max_column)  # (2 / 7 / 12)
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                # Program Property
                program_property @title
                @program.adjustments.destroy_all
              end
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row +1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*(c_i)
                        @program.save
                      end
                      @block_hash[main_key][key][15*(c_i)] = value unless @block_hash[main_key][key].nil?
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

        # VA loan level adjustment
        (15..33).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (17..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "VA Loan Level Adjustments"
                  primary_key = "VA/RateType/FICO/LTV"
                  @valn_adjustment[primary_key] = {}
                end

                if r >=17 && r <= 31 && cc == 17
                  value = get_value value
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @valn_adjustment[primary_key][value] = c_val
                end

                if r >=32 && r <= 33 && cc == 17
                  value = get_value value
                  ccc = cc + 2
                  c_val = sheet_data.cell(r,ccc)
                  @valn_adjustment[primary_key][value] = c_val
                end
              end
            end
          end
        end

        # WholeSale adjustment
        (34..45).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (17..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Allied Wholesale Loan Amt Adj *"
                  primary_key = "VA/RateType/FICO/LTV"
                  @awlm_adjustment[primary_key] = {}
                end

                if r >=37 && r <= 45 && cc == 17
                  value = get_value value
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @awlm_adjustment[primary_key][value] = c_val
                end
              end
            end
          end
        end

        # INDICES adjustment
        (49..54).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (18..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "INDICES"
                  primary_key = "VA/RateType/FICO/LTV"
                  second_key = "INDICES"
                  @indi_adjustment[primary_key] = {}
                  @indi_adjustment[primary_key][second_key] = {}
                end

                if r >=50 && r <= 54 && cc == 18
                  value = get_value value
                  ccc = cc + 2
                  c_val = sheet_data.cell(r,ccc)
                  @indi_adjustment[primary_key][second_key][value] = c_val
                end
              end
            end
          end
        end

        #Adjustments Avrage Price
        (88..90).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (2..9).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Average Price Offer Rate (APOR) For Loans Locked During The Week Of 01/14/2019"
                  f_key = "AveragePrice"
                  @avrg_adjustment[f_key] = {}
                end

                if r == 89 && cc >= 2 && cc <= 9
                  rr = r+1
                  c_val = sheet_data.cell(rr,cc)
                  @avrg_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end

        #Adjustment lock Expiration
        (83..86).each do |r|
          row = sheet_data.row(r)
          @key_data = sheet_data.row(93)
          @key_data2 = sheet_data.row(93)
          if (row.compact.count >= 1)
            (2..5).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Lock Expiration Dates:"
                  f_key = "Lock/ExpirationDates:"
                  @lock_adjustment[f_key] = {}
                end

                if r >=84 && r <= 86 && cc == 2
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @lock_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end
        Adjustment.create(data: @valn_adjustment, sheet_name: sheet)
        Adjustment.create(data: @awlm_adjustment, sheet_name: sheet)
        Adjustment.create(data: @indi_adjustment, sheet_name: sheet)
        Adjustment.create(data: @lock_adjustment, sheet_name: sheet)
        create_program_association_with_adjustment(sheet)
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
        @valn_adjustment = {}
        @avrg_adjustment = {}
        @lock_adjustment = {}
        @ltv_fico_hash = {}
        @sub_adjustment = {}
        @subordinate_hash = {}
        @home_hash = {}
        @ltv_co_hash = {}
        @occp_hash = {}
        @pro_hash ={}
        @key_data = []
        @cltv_data = []
        @programs_ids = []
        primary_key = ''
        second_key = ''
        ltv_key = ''
        cltv_key = ''
        k_value = ''
        c_val = ''
        f_key = ''
        value1 = ''
        value2 = ''
        value3 = ''

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
              if @program.term.present?
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
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
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        # VA loan level adjustment
        (88..102).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (13..16).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Allied Wholesale Loan Amt Adj"
                  primary_key = "AlliedWholesale/LoanAmount/"

                  @valn_adjustment[primary_key] = {}
                end

                if r >=91 && r <= 99 && cc == 13
                  value = get_value value
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @valn_adjustment[primary_key][value] = c_val
                end

                if r >= 101 && r <= 102 && cc == 13
                  value = get_value value
                  ccc = cc + 2
                  c_val = sheet_data.cell(r,ccc)
                  @valn_adjustment[primary_key][value] = c_val
                end
              end
            end
          end
        end

        #Adjustments Avrage Price
        (109..111).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (7..14).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Average Price Offer Rate (APOR) For Loans Locked During The Week Of 01/14/2019"
                  f_key = "AveragePrice"
                  @avrg_adjustment[f_key] = {}
                end

                if r == 110 && cc >= 7 && cc <= 14
                  rr = r+1
                  c_val = sheet_data.cell(rr,cc)
                  @avrg_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end

        #Adjustment lock Expiration
        (109..112).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (2..5).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Lock Expiration Dates:"
                  f_key = "Lock/ExpirationDates:"
                  @lock_adjustment[f_key] = {}
                end

                if r >=110 && r <= 112 && cc == 2
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @lock_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end

        #Adjustment LTV/FICO
        (63..83).each do |r|
          row = sheet_data.row(r)
          @key_data = sheet_data.row(63)
          if (row.compact.count >= 1)
            (2..12).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "LTV"
                  f_key = "LoanType/Term/FICO/LTV"
                  @ltv_fico_hash[f_key] = {}
                end

                if r >=65 && r <= 71 && cc == 4
                  value1 = get_value value
                  @ltv_fico_hash[f_key][value1] = {}
                end

                if r >=65 && r <= 71 && cc >= 5 && cc<= 12
                  k_value = get_value @key_data[cc-2]
                  @ltv_fico_hash[f_key][value1][k_value] = value
                end

                if value == "Cash-Out"
                  primary_key = "LoanType/Term/FICO/LTV"
                  @ltv_co_hash[primary_key] ={}
                end

                if r >=72 && r <= 78 && cc == 4
                  value2 = get_value value
                  @ltv_co_hash[primary_key][value2] = {}
                end

                if r >=72 && r <= 78 && cc >= 5 && cc<= 12
                  k_value = get_value @key_data[cc-2]
                  @ltv_co_hash[primary_key][value2][k_value] = value
                end

                if value == "OCCUPANCY"
                  primary_key = "LoanType/Term/FICO/LTV"
                  @occp_hash[primary_key] ={}
                end

                if r >=79 && r <= 80 && cc == 4
                  value3 = get_value value
                  @occp_hash[primary_key][value3] = {}
                end

                if r >=79 && r <= 80 && cc >= 5 && cc<= 12
                  k_value = get_value @key_data[cc-2]
                  @occp_hash[primary_key][value3][k_value] = value
                end

                if value == "Property Type"
                  primary_key = "LoanType/Term/FICO/LTV"
                  @pro_hash[primary_key] ={}
                end

                if r >=81 && r <= 83 && cc == 4
                  value3 = get_value value
                  @pro_hash[primary_key][value3] = {}
                end

                if r >=81 && r <= 83 && cc >= 5 && cc<= 12
                  k_value = get_value @key_data[cc-2]
                  @pro_hash[primary_key][value3][k_value] = value
                end

              end
            end
          end
        end

        #Adjustment Subordinate Finance
        (63..77).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(64)
          if (row.compact.count >= 1)
            (13..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "SUBORDINATE FINANCING"
                  f_key = "LTV/CLTV/FICO"
                  primary_key = "LTV/CLTV/FICO"
                  @sub_adjustment[f_key] = {}
                  @subordinate_hash[primary_key] = {}
                end

                if r >=70 && r <= 77 && cc == 13
                  value1 = get_value value
                  ccc = cc + 7
                  c_val = sheet_data.cell(r,ccc)
                  @sub_adjustment[f_key][value] = c_val
                end

                # SUBORDINATE FINANCING
                if r >= 65 && r <= 69 && cc == 13
                  second_key = get_value value
                  @subordinate_hash[primary_key][second_key] = {}
                end
                if r >= 65 && r <= 69 && cc == 15
                  ltv_key = get_value value
                  @subordinate_hash[primary_key][second_key][ltv_key] = {}
                end
                if r >= 65 && r <= 69 && cc >= 17 && cc <= 20
                  cltv_key = @cltv_data[cc-2]
                  @subordinate_hash[primary_key][second_key][ltv_key] = {}
                  @subordinate_hash[primary_key][second_key][ltv_key][cltv_key] = value
                end
              end
            end
          end
        end

        #Adjustment HomeReady
        (82..84).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (13..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "HomeReady/Home Possible LLPA Caps"
                  f_key = "LoanType/Term/FICO/LTV"
                  @home_hash[f_key] = {}
                end

                if r >=83 && r <= 84 && cc == 13
                  ccc = cc + 7
                  c_val = sheet_data.cell(r,ccc)
                  @home_hash[f_key][value] = c_val
                end

              end
            end
          end
        end
        Adjustment.create(data: @home_hash, sheet_name: sheet)
        Adjustment.create(data: @subordinate_hash, sheet_name: sheet)
        Adjustment.create(data: @sub_adjustment, sheet_name: sheet)
        Adjustment.create(data: @occp_hash, sheet_name: sheet)
        Adjustment.create(data: @ltv_co_hash, sheet_name: sheet)
        Adjustment.create(data: @ltv_fico_hash, sheet_name: sheet)
        Adjustment.create(data: @valn_adjustment, sheet_name: sheet)
        Adjustment.create(data: @avrg_adjustment, sheet_name: sheet)
        Adjustment.create(data: @lock_adjustment, sheet_name: sheet)
        create_program_association_with_adjustment(sheet)
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
        if value1.include?("FICO <")
          value1 = "0"+value1.split("FICO").last
        elsif value1.include?("<=") || value1.include?(">=")
          value1 = "0"+value1
        elsif value1.include?("FICO")
          value1 = value1.split("FICO ").last.first(9)
        elsif value1 == "Investment Property"
          value1 = "Property/Type"
        else
          value1
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
      if @program.program_name.include?("High Bal")
        @jumbo_high_balance = true
      end

       # Program Category
      if @program.program_name.include?("F30/F25")
        @program_category = "F30/F25"
      elsif @program.program_name.include?("F15")
        @program_category = "F15"
      elsif @program.program_name.include?("f30J")
        @program_category = "f30J"
      elsif @program.program_name.include?("F15S")
        @program_category = "F15S"
      elsif @program.program_name.include?("F30JS")
        @program_category = "F30JS"
      elsif @program.program_name.include?("F5YT")
        @program_category = "F5YT"
      elsif @program.program_name.include?("F5YTS")
        @program_category = "F5YTS"
      elsif @program.program_name.include?("F5YTJ")
        @program_category = "F5YTJ"
      end

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
      @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline)
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
