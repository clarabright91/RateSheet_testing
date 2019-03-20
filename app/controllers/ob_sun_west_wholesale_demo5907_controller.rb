class ObSunWestWholesaleDemo5907Controller < ApplicationController
  before_action :get_sheet, only: [:programs, :ratesheet]
  before_action :get_program, only: [:single_program]
  def index
    file = File.join(Rails.root,  'OB_SunWest_Wholesale_Demo5907.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "RATESHEET")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "SunWest Wholesale"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def ratesheet
    file = File.join(Rails.root,  'OB_SunWest_Wholesale_Demo5907.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "RATESHEET")
        sheet_data = xlsx.sheet(sheet)
        @adj_hash = {}
        @property_hash = {}
        @jumbo_hash = {}
        @price_hash = {}
        @gov_hash = {}
        primary_key = ''
        ltv_key = ''
        cltv_key = ''
        first_key = ''
        second_key = ''
        c_val = ''
        @ltv_data = []

        @programs_ids = []
        first_key = []
        @key_data = []
        @conf_adjustment = {}
        k_value = ''
        value1 = ''
        range1 = 374
        range2 = 404
        # range1_a = 782
        # range2_a = 799
        # range1_b = 1100
        # range2_b = 1199
        # range1_c = 1599
        # range2_c = 1614
        # range1_d = 2434
        # range2_d = 2544
        # range1_e = 2624
        # range2_e = 2738
        # range1_f = 2791
        # range2_f = 2885
        # range1_g = 2936
        # range2_g = 3036
        # range1_h = 3089
        # range2_h = 3183
        # range1_i = 3237
        # range2_i = 3280
        # range1_j = 3334
        # range2_j = 3379
        # range1_k = 3433
        # range2_k = 3527

        # Agency Conforming Programs
        (156..320).each do |r|
          row = sheet_data.row(r)
          @sheet_name = "AGENCY CONFORMING PROGRAMS"
          @bank = @sheet_obj.bank
          @sheet_obj = @bank.sheets.find_or_create_by(name: @sheet_name)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + 2 # 2 / 7 / 12 / 17
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "Rate"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                end

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
                        @block_hash[key][15*(c_i+1)] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil? || @block_hash.keys.first == "Rate"
                  @block_hash.shift
                end
                @program.update(base_rate: @block_hash)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # PRICE ADJUSTMENTS: CONFORMING PROGRAMS //adjustment
        (374..404).each do |r|
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "LOAN TERM > 15 YEARS"
                  primary_key = "LoanSize/Term/FICO/LTV"
                  first_row = 377
                  end_row = 384
                  last_column = 10
                  first_column = 2
                  ltv_row = 375
                  ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
                end
                if value == "CASH OUT REFINANCE "
                  primary_key = "LoanSize/RefinanceOption/FICO/LTV"
                  first_row = 389
                  end_row = 396
                  first_column = 2
                  last_column = 6
                  ltv_row = 387
                  cash_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
                end
                if value == "ADDITIONAL LPMI ADJUSTMENTS"
                  primary_key = "LPMI/LoanSize/FICO/LTV"
                  first_row = 390
                  end_row = 393
                  first_column = 9
                  last_column = 12
                  ltv_row = 388
                  lpmi_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
                end
                if value == "SUBORDINATE FINANCING"
                  primary_key = "FinancingType/LoanSize/LTV/CLTV/FICO"
                  first_row = 400
                  end_row = 404
                  first_column = 2
                  cltv_column = 4
                  last_column = 7
                  ltv_row = 399
                  sub_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, cltv_column, last_column, ltv_row, primary_key
                end
                if value == "LPMI COVERAGE BASED ADJUSTMENTS"
                  primary_key = "LPMI/RefinanceOption/FICO"
                  first_row = 399
                  end_row = 404
                  first_column = 9
                  last_column = 12
                  ltv_row = 397
                  lpmi_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
                end
                if r == 376 && cc == 15
                  @adj_hash["PropertyType"] = {}
                  @adj_hash["PropertyType"]["2-4 Unit"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["PropertyType"]["2-4 Unit"] = new_val
                end

                if r == 377 && cc == 15
                  @adj_hash["PropertyType/Term/LTV"] = {}
                  @adj_hash["PropertyType/Term/LTV"]["Condo"] = {}
                  @adj_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                  @adj_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = new_val
                end
                if r == 378 && cc == 15
                  @adj_hash["PropertyType"]["Manufactured Home"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["PropertyType"]["Manufactured Home"] = new_val
                end
                if r == 379 && cc == 15
                  @adj_hash["FinancingType"] = {}
                  @adj_hash["FinancingType"]["Subordinate Financing"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["FinancingType"]["Subordinate Financing"] = new_val
                end
                if r == 380 && cc == 15
                  @adj_hash["PropertyType/LTV"] = {}
                  @adj_hash["PropertyType/LTV"]["Investment Property"] = {}
                  @adj_hash["PropertyType/LTV"]["Investment Property"]["0-75"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["PropertyType/LTV"]["Investment Property"]["0-75"] = new_val
                end
                if r == 381 && cc == 15
                  @adj_hash["PropertyType/LTV"]["Investment Property"]["75.01-80.01"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["PropertyType/LTV"]["Investment Property"]["75.01-80.01"] = new_val
                end
                if r == 382 && cc == 15
                  @adj_hash["PropertyType/LTV"]["Investment Property"]["80.01-85.00"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["PropertyType/LTV"]["Investment Property"]["80.01-85.00"] = new_val
                end
                if r == 383 && cc == 15
                  @adj_hash["LoanType/LTV"] = {}
                  @adj_hash["LoanType/LTV"]["ARM"] = {}
                  @adj_hash["LoanType/LTV"]["ARM"]["90-Inf"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["LoanType/LTV"]["ARM"]["90-Inf"] = new_val
                end
                if r == 384 && cc == 15
                  @adj_hash["LPMI/RefinanceOption/FICO"] = {}
                  @adj_hash["LPMI/RefinanceOption/FICO"][true] = {}
                  @adj_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"] = {}
                  @adj_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"]["680-719"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"]["680-719"] = new_val
                end
                if r == 385 && cc == 15
                  @adj_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"]["660-679"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"]["660-679"] = new_val
                end
                if r == 386 && cc == 15
                  @adj_hash["LoanSize/LoanType/LTV"] = {}
                  @adj_hash["LoanSize/LoanType/LTV"]["High-Balance"] = {}
                  @adj_hash["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"] = {}
                  @adj_hash["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["0-75"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["0-75"] = new_val
                end
                if r == 387 && cc == 15
                  @adj_hash["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["75-Inf"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["75-Inf"] = new_val
                end
                if r == 388 && cc == 15
                  @adj_hash["LoanSize/RefinanceOption"] = {}
                  @adj_hash["LoanSize/RefinanceOption"]["High-Balance"] = {}
                  @adj_hash["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = new_val
                end
                if r == 389 && cc == 15
                  @adj_hash["LoanPurpose/LoanSize/RefinanceOption"] = {}
                  @adj_hash["LoanPurpose/LoanSize/RefinanceOption"]["Purchase"] = {}
                  @adj_hash["LoanPurpose/LoanSize/RefinanceOption"]["Purchase"]["High-Balance"] = {}
                  @adj_hash["LoanPurpose/LoanSize/RefinanceOption"]["Purchase"]["High-Balance"]["Cash Out"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["LoanPurpose/LoanSize/RefinanceOption"]["Purchase"]["High-Balance"]["Cash Out"] = new_val
                end
                if r == 397 && cc == 15
                  @adj_hash["FNMA/FannieMaeProduct/FICO/LTV"] = {}
                  @adj_hash["FNMA/FannieMaeProduct/FICO/LTV"][true] = {}
                  @adj_hash["FNMA/FannieMaeProduct/FICO/LTV"][true]["HomeReady"] = {}
                  @adj_hash["FNMA/FannieMaeProduct/FICO/LTV"][true]["HomeReady"]["680-Inf"] = {}
                  @adj_hash["FNMA/FannieMaeProduct/FICO/LTV"][true]["HomeReady"]["680-Inf"]["80-Inf"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["FNMA/FannieMaeProduct/FICO/LTV"][true]["HomeReady"]["680-Inf"]["80-Inf"] = new_val
                end
                if r == 398 && cc == 15
                  @adj_hash["FNMA/FannieMaeProduct/FICO/LTV"][true]["HomeReady"]["0-680"] = {}
                  @adj_hash["FNMA/FannieMaeProduct/FICO/LTV"][true]["HomeReady"]["0-680"]["0-80"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @adj_hash["FNMA/FannieMaeProduct/FICO/LTV"][true]["HomeReady"]["0-680"]["0-80"] = new_val
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end
        adjustment = [@adj_hash]
        make_adjust(adjustment,@sheet_name)

        # FHLMC HOME POSSIBLE / HOMEONE / SUPER CONFORMING Programs
        (708..760).each do |r|
          row = sheet_data.row(r)
          @sheet_name = "FHLMC HOME POSSIBLE / HOMEONE / SUPER CONFORMING"
          @bank = @sheet_obj.bank
          @sheet_obj = @bank.sheets.find_or_create_by(name: @sheet_name)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + 2 # 2 / 7 / 12 / 17
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "Rate"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                end

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
                        @block_hash[key][15*(c_i+1)] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil? || @block_hash.keys.first == "Rate"
                  @block_hash.shift
                end
                @program.update(base_rate: @block_hash)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # PRICE ADJUSTMENTS: FHLMC HOME POSSIBLE / HOMEONE / SUPER CONFORMING //adjustment // 3 more adjustment remaining for this programs
        (782..799).each do |r|
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "LOAN TERM > 15 YEARS"
                  primary_key = "LoanSize/Term/FICO/LTV"
                  first_row = 785
                  end_row = 791
                  first_column = 2
                  last_column = 11
                  ltv_row = 783
                  start_range = 782
                  end_range = 799
                  ltv_adjustment start_range, end_range, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
                end
                if value == "SUBORDINATE FINANCING  (Applicable to HomeOne)"
                  primary_key = "FinancingType/LoanSize/LTV/CLTV/FICO"
                  first_row = 795
                  end_row = 799
                  first_column = 15
                  cltv_column = 17
                  last_column = 20
                  ltv_row = 794
                  range1 = 782
                  range2 = 799
                  sub_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, cltv_column, last_column, ltv_row, primary_key
                end
                if r == 783 && cc == 15
                  @property_hash["PropertyType/Term/LTV"] = {}
                  @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                  @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                  @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = new_val
                end
                if r == 784 && cc == 15
                  @property_hash["PropertyType"] = {}
                  @property_hash["PropertyType"]["Manufactured Home"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @property_hash["PropertyType"]["Manufactured Home"] = new_val
                end
                if r == 785 && cc == 15
                  @property_hash["PropertyType"]["2 Unit"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @property_hash["PropertyType"]["2 Unit"] = new_val
                end
                if r == 786 && cc == 15
                  @property_hash["PropertyType/LTV"] = {}
                  @property_hash["PropertyType/LTV"]["3-4 Unit"] = {}
                  @property_hash["PropertyType/LTV"]["3-4 Unit"]["0-80"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @property_hash["PropertyType/LTV"]["3-4 Unit"]["0-80"] = new_val
                end
                if r == 787 && cc == 15
                  @property_hash["PropertyType/LTV"]["3-4 Unit"]["80-85"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @property_hash["PropertyType/LTV"]["3-4 Unit"]["80-85"] = new_val
                end
                if r == 788 && cc == 15
                  @property_hash["PropertyType/LTV"]["3-4 Unit"]["85-Inf"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @property_hash["PropertyType/LTV"]["3-4 Unit"]["85-Inf"] = new_val
                end
                if r == 789 && cc == 15
                  @property_hash["FinancingType"] = {}
                  @property_hash["FinancingType"]["Subordinate Financing"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @property_hash["FinancingType"]["Subordinate Financing"] = new_val
                end
                if r == 791 && cc == 15
                  @property_hash["FreddieMacProduct/LoanPurpose/LoanSize/RefinanceOption"] = {}
                  @property_hash["FreddieMacProduct/LoanPurpose/LoanSize/RefinanceOption"]["HomePossible"] = {}
                  @property_hash["FreddieMacProduct/LoanPurpose/LoanSize/RefinanceOption"]["HomePossible"]["Purchase"] = {}
                  @property_hash["FreddieMacProduct/LoanPurpose/LoanSize/RefinanceOption"]["HomePossible"]["Purchase"]["Conforming"] = {}
                  @property_hash["FreddieMacProduct/LoanPurpose/LoanSize/RefinanceOption"]["HomePossible"]["Purchase"]["Conforming"]["Cash Out"] = {}
                  cc = cc + 5
                  new_val = sheet_data.cell(r,cc)
                  @property_hash["FreddieMacProduct/LoanPurpose/LoanSize/RefinanceOption"]["HomePossible"]["Purchase"]["Conforming"]["Cash Out"] = new_val
                end
                if r == 795 && cc == 2
                  @property_hash["FreddieMacProduct/LTV/FICO"] = {}
                  @property_hash["FreddieMacProduct/LTV/FICO"]["HomePossible"] = {}
                  @property_hash["FreddieMacProduct/LTV/FICO"]["HomePossible"]["80-Inf"] = {}
                  @property_hash["FreddieMacProduct/LTV/FICO"]["HomePossible"]["80-Inf"]["680-Inf"] = {}
                  cc = cc + 2
                  new_val = sheet_data.cell(r,cc)
                  @property_hash["FreddieMacProduct/LTV/FICO"]["HomePossible"]["80-Inf"]["680-Inf"] = new_val
                end
                if r == 796 && cc == 3
                  @property_hash["FreddieMacProduct/LTV/FICO"]["HomePossible"]["80-Inf"]["0-680"] = {}
                  cc = cc + 1
                  new_val = sheet_data.cell(r,cc)
                  @property_hash["FreddieMacProduct/LTV/FICO"]["HomePossible"]["80-Inf"]["0-680"] = new_val
                end
                if r == 797 && cc == 2
                  @property_hash["FreddieMacProduct/LTV/FICO"]["HomePossible"]["0-80"] = {}
                  @property_hash["FreddieMacProduct/LTV/FICO"]["HomePossible"]["0-80"]["0-Inf"] = {}
                  cc = cc + 2
                  new_val = sheet_data.cell(r,cc)
                  @property_hash["FreddieMacProduct/LTV/FICO"]["HomePossible"]["0-80"]["0-Inf"] = new_val
                end
              end
              #ADJUSTMENT CAPS (Applicable to Home Possible products) Not completed
              if value == "ADJUSTMENT CAPS (Applicable to Home Possible products)"
                primary_key = "LoanType/Term/LTV/FICO"
                @caps_adjustment[primary_key] = {}
              end
              # ADJUSTMENT CAPS end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end
        adjustment = [@property_hash]
        make_adjust(adjustment,@sheet_name)

        #Non-Confirming: Sigma Programs
        (1101..1179).each do |r|
          row = sheet_data.row(r)
          @sheet_name = "NON-CONFORMING: SIGMA QM PRIME JUMBO"
          @bank = @sheet_obj.bank
          @sheet_obj = @bank.sheets.find_or_create_by(name: @sheet_name)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + 2 # 2 / 7 / 12 / 17
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "ARM INFORMATION"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                end
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
                        @block_hash[key][15*(c_i+1)] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil? || @block_hash.keys.first == "Rate"
                  @block_hash.shift
                end
                @program.update(base_rate: @block_hash)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #NON-CONFORMING: SIGMA QM PRIME JUMBO //adjustment
        (1166..1199).each do |r|
          @jumbo_data = sheet_data.row(1167)
          @price_data = sheet_data.row(1183)
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "PROGRAM SPECIFIC PRICE ADJUSTMENTS"
                  @jumbo_hash["State/LTV"] = {}
                  @jumbo_hash["LoanAmount/FICO/LTV"] = {}
                  @jumbo_hash["LoanAmount/FICO/LTV"]["0-1,000,000"] = {}
                  @jumbo_hash["LoanAmount/FICO/LTV"]["1,000,000-1,500,000"] = {}
                  @jumbo_hash["LoanAmount/FICO/LTV"]["1,500,000-2,000,000"] = {}
                  @jumbo_hash["LoanAmount/FICO/LTV"]["2,000,000-2,500,000"] = {}
                end
                if r == 1169 && cc == 14
                  @jumbo_hash["LockDay/LTV"] = {}
                  @jumbo_hash["LockDay/LTV"]["15"] = {}
                end
                if r == 1169 && cc >= 16 && cc <= 20
                  ltv_key = get_value @jumbo_data[cc-2]
                  @jumbo_hash["LockDay/LTV"]["15"][ltv_key] = {}
                  @jumbo_hash["LockDay/LTV"]["15"][ltv_key] = value
                end
                if r >= 1170 && r <= 1174 && cc == 14
                  primary_key = value.split.last
                  @jumbo_hash["State/LTV"][primary_key] = {}
                end
                if r >= 1170 && r <= 1174 && cc >= 16 && cc <= 20
                  ltv_key = get_value @jumbo_data[cc-2]
                  @jumbo_hash["State/LTV"][primary_key][ltv_key] = {}
                  @jumbo_hash["State/LTV"][primary_key][ltv_key] = value
                end
                if r == 1175 && cc == 14
                  @jumbo_hash["PropertyType/LTV"] = {}
                  @jumbo_hash["PropertyType/LTV"]["2 Unit"] = {}
                end
                if r == 1175 && cc >= 16 && cc <= 20
                  ltv_key = get_value @jumbo_data[cc-2]
                  @jumbo_hash["PropertyType/LTV"]["2 Unit"][ltv_key] = {}
                  @jumbo_hash["PropertyType/LTV"]["2 Unit"][ltv_key] = value
                end
                if r == 1176 && cc == 14
                  @jumbo_hash["PropertyType/LTV"]["Condo"] = {}
                end
                if r == 1176 && cc >= 16 && cc <= 20
                  ltv_key = get_value @jumbo_data[cc-2]
                  @jumbo_hash["PropertyType/LTV"]["Condo"][ltv_key] = {}
                  @jumbo_hash["PropertyType/LTV"]["Condo"][ltv_key] = value
                end
                if r == 1177 && cc == 14
                  @jumbo_hash["PropertyType/LTV"]["2nd Home"] = {}
                end
                if r == 1177 && cc >= 16 && cc <= 20
                  ltv_key = get_value @jumbo_data[cc-2]
                  @jumbo_hash["PropertyType/LTV"]["2nd Home"][ltv_key] = {}
                  @jumbo_hash["PropertyType/LTV"]["2nd Home"][ltv_key] = value
                end
                if r == 1178 && cc == 14
                  @jumbo_hash["LoanPurpose/LTV"] = {}
                  @jumbo_hash["LoanPurpose/LTV"]["Purchase"] = {}
                end
                if r == 1178 && cc >= 16 && cc <= 20
                  ltv_key = get_value @jumbo_data[cc-2]
                  @jumbo_hash["LoanPurpose/LTV"]["Purchase"][ltv_key] = {}
                  @jumbo_hash["LoanPurpose/LTV"]["Purchase"][ltv_key] = value
                end
                if r == 1179 && cc == 14
                  @jumbo_hash["RefinanceOption/LTV"] = {}
                  @jumbo_hash["RefinanceOption/LTV"]["Cash Out"] = {}
                end
                if r == 1179 && cc >= 16 && cc <= 20
                  ltv_key = get_value @jumbo_data[cc-2]
                  @jumbo_hash["RefinanceOption/LTV"]["Cash Out"][ltv_key] = {}
                  @jumbo_hash["RefinanceOption/LTV"]["Cash Out"][ltv_key] = value
                end
                if r >= 1185 && r <= 1189 && cc == 2
                  primary_key = get_value value
                  @jumbo_hash["LoanAmount/FICO/LTV"]["0-1,000,000"][primary_key] = {}
                  @jumbo_hash["LoanAmount/FICO/LTV"]["1,000,000-1,500,000"][primary_key] = {}
                end
                if r >= 1185 && r <= 1189 && cc >= 3 && cc <= 7
                  ltv_key = get_value @price_data[cc-2]
                  @jumbo_hash["LoanAmount/FICO/LTV"]["0-1,000,000"][primary_key][ltv_key] = {}
                  @jumbo_hash["LoanAmount/FICO/LTV"]["0-1,000,000"][primary_key][ltv_key] = value
                end
                if r >= 1185 && r <= 1189 && cc >= 8 && cc <= 12
                  ltv_key = get_value @price_data[cc-2]
                  @jumbo_hash["LoanAmount/FICO/LTV"]["1,000,000-1,500,000"][primary_key][ltv_key] = {}
                  @jumbo_hash["LoanAmount/FICO/LTV"]["1,000,000-1,500,000"][primary_key][ltv_key] = value
                end
                if r >= 1195 && r <= 1198 && cc == 2
                  primary_key = get_value value
                  @jumbo_hash["LoanAmount/FICO/LTV"]["1,500,000-2,000,000"][primary_key] = {}
                  @jumbo_hash["LoanAmount/FICO/LTV"]["2,000,000-2,500,000"][primary_key] = {}
                end
                if r >= 1195 && r <= 1198 && cc >= 3 && cc <= 7
                  ltv_key = get_value @price_data[cc-2]
                  @jumbo_hash["LoanAmount/FICO/LTV"]["1,500,000-2,000,000"][primary_key][ltv_key] = {}
                  @jumbo_hash["LoanAmount/FICO/LTV"]["1,500,000-2,000,000"][primary_key][ltv_key] = value
                end
                if r >= 1195 && r <= 1198 && cc >= 8 && cc <= 12
                  ltv_key = get_value @price_data[cc-2]
                  @jumbo_hash["LoanAmount/FICO/LTV"]["2,000,000-2,500,000"][primary_key][ltv_key] = {}
                  @jumbo_hash["LoanAmount/FICO/LTV"]["2,000,000-2,500,000"][primary_key][ltv_key] = value
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end
        adjustment = [@jumbo_hash]
        make_adjust(adjustment,@sheet_name)
        #NON-CONFORMING: JW /Programs
        (1386..1547).each do |r|
          row = sheet_data.row(r)
          @sheet_name = "NON-CONFORMING: JW"
          @bank = @sheet_obj.bank
          @sheet_obj = @bank.sheets.find_or_create_by(name: @sheet_name)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + 2 # 2 / 7 / 12 / 17
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "Rate"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                end

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
                        @block_hash[key][15*(c_i+1)] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil? || @block_hash.keys.first == "Rate"
                  @block_hash.shift
                end
                @program.update(base_rate: @block_hash)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #NON-CONFORMING: JW /adjustment
        (1599..1614).each do |r|
          @price_data = sheet_data.row(1600)
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "PRICE ADJUSTMENTS"
                  @price_hash["LoanSize/FICO/LTV"] = {}
                  @price_hash["LoanSize/FICO/LTV"]["Non-Conforming"] = {}
                end
                if value == "STATE SPECEFIC PRICE ADJUSTMENTS"
                  @price_hash["LoanSize/LoanType/State/Term"] = {}
                  @price_hash["LoanSize/LoanType/State/Term"]["Non-Conforming"] = {}
                  @price_hash["LoanSize/LoanType/State/Term"]["Non-Conforming"]["Fixed"] = {}
                  @price_hash["LoanSize/LoanType/State/ArmBasic"] = {}
                  @price_hash["LoanSize/LoanType/State/ArmBasic"]["Non-Conforming"] = {}
                  @price_hash["LoanSize/LoanType/State/ArmBasic"]["Non-Conforming"]["ARM"] = {}
                end
                if value == "PROGRAM SPECIFIC PRICE ADJUSTMENTS"
                  @price_hash["LoanSize/RefinanceOption/LTV"] = {}
                  @price_hash["LoanSize/RefinanceOption/LTV"]["Non-Conforming"] = {}
                  @price_hash["LoanSize/RefinanceOption/LTV"]["Non-Conforming"]["Cash Out"] = {}
                end
                if r >= 1602 && r <= 1607 && cc == 7
                  primary_key = get_value value
                  @price_hash["LoanSize/FICO/LTV"]["Non-Conforming"][primary_key] = {}
                end
                if r >= 1602 && r <= 1607 && cc >= 8 && cc <= 11
                  ltv_key = get_value @price_data[cc-2]
                  @price_hash["LoanSize/FICO/LTV"]["Non-Conforming"][primary_key][ltv_key] = {}
                  @price_hash["LoanSize/FICO/LTV"]["Non-Conforming"][primary_key][ltv_key] = value
                end
                if r >= 1601 && r <= 1610 && cc == 13
                  cltv_key = value.split.last
                  @price_hash["LoanSize/LoanType/State/Term"]["Non-Conforming"]["Fixed"][cltv_key] = {}
                  @price_hash["LoanSize/LoanType/State/ArmBasic"]["Non-Conforming"]["ARM"][cltv_key] = {}
                end
                if r >= 1601 && r <= 1610 && cc >= 15 && cc <= 17
                  first_key = get_value @price_data[cc-2]
                  @price_hash["LoanSize/LoanType/State/Term"]["Non-Conforming"]["Fixed"][cltv_key][first_key] = {}
                  @price_hash["LoanSize/LoanType/State/Term"]["Non-Conforming"]["Fixed"][cltv_key][first_key] = value
                end
                if r >= 1601 && r <= 1610 && cc >= 18 && cc <= 20
                  first_key = @price_data[cc-2].split("/").first
                  @price_hash["LoanSize/LoanType/State/ArmBasic"]["Non-Conforming"]["ARM"][cltv_key][first_key] = {}
                  @price_hash["LoanSize/LoanType/State/ArmBasic"]["Non-Conforming"]["ARM"][cltv_key][first_key] = value
                end
                if r == 1607 && cc == 2
                  @price_hash["LoanSize/PropertyType"] = {}
                  @price_hash["LoanSize/PropertyType"]["Non-Conforming"] = {}
                  @price_hash["LoanSize/PropertyType"]["Non-Conforming"]["Investment Property"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @price_hash["LoanSize/PropertyType"]["Non-Conforming"]["Investment Property"] = new_val
                end
                if r == 1608 && cc == 2
                  @price_hash["LoanSize/PropertyType"]["Non-Conforming"]["3 Unit"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @price_hash["LoanSize/PropertyType"]["Non-Conforming"]["3 Unit"] = new_val
                end
                if r == 1609 && cc == 2
                  @price_hash["LoanSize/PropertyType"]["Non-Conforming"]["4 Unit"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @price_hash["LoanSize/PropertyType"]["Non-Conforming"]["4 Unit"] = new_val
                end
                if r == 1610 && cc == 2
                  @price_hash["LoanSize/PropertyType"]["Non-Conforming"]["2nd Home"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @price_hash["LoanSize/PropertyType"]["Non-Conforming"]["2nd Home"] = new_val
                end
                if r == 1611 && cc == 2
                  @price_hash["LoanSize/LoanAmount"] = {}
                  @price_hash["LoanSize/LoanAmount"]["Non-Conforming"] = {}
                  @price_hash["LoanSize/LoanAmount"]["Non-Conforming"]["1,000,000-Inf"] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @price_hash["LoanSize/LoanAmount"]["Non-Conforming"]["1,000,000-Inf"] = new_val
                end
                if r >= 1612 && r <= 1614 && cc == 2
                  ltv_key = get_value value
                  @price_hash["LoanSize/RefinanceOption/LTV"]["Non-Conforming"]["Cash Out"][ltv_key] = {}
                  cc = cc + 3
                  new_val = sheet_data.cell(r,cc)
                  @price_hash["LoanSize/RefinanceOption/LTV"]["Non-Conforming"]["Cash Out"][ltv_key] = new_val
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end
        adjustment = [@price_hash]
        make_adjust(adjustment,@sheet_name)

        #GOVERNMENT PROGRAMS /programs
        (2180..2278).each do |r|
          row = sheet_data.row(r)
          @sheet_name = "GOVERNMENT PROGRAMS"
          @bank = @sheet_obj.bank
          @sheet_obj = @bank.sheets.find_or_create_by(name: @sheet_name)
          if ((row.compact.count >= 3) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + 2 # 2 / 7 / 12 / 17
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "Rate" && @title != "PROGRAM SPECIFIC PRICE ADJUSTMENTS"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                end

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
                        @block_hash[key][15*(c_i+1)] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil? || @block_hash.keys.first == "Rate"
                  @block_hash.shift
                end
                @program.update(base_rate: @block_hash)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #GOVERNMENT PROGRAMS /adjustments
        (2255..2278).each do |r|
          @ltv_data = sheet_data.row(2274)
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "PROGRAM SPECIFIC PRICE ADJUSTMENTS"
                  @gov_hash["FICO"] = {}
                end
                if r == 2256 && cc == 14
                  @gov_hash["FHA"] = {}
                  @gov_hash["FHA"][true] = {}
                  cc = cc + 6
                  new_val = sheet_data.cell(r,cc)
                  @gov_hash["FHA"][true] = new_val
                end
                if r == 2257 && cc == 14
                  @gov_hash["VA"] = {}
                  @gov_hash["VA"][true] = {}
                  cc = cc + 6
                  new_val = sheet_data.cell(r,cc)
                  @gov_hash["VA"][true] = new_val
                end
                if r == 2258 && cc == 14
                  @gov_hash["USDA"] = {}
                  @gov_hash["USDA"][true] = {}
                  cc = cc + 6
                  new_val = sheet_data.cell(r,cc)
                  @gov_hash["USDA"][true] = new_val
                end
                if r == 2260 && cc == 14
                  @gov_hash["LoanType/Term"] = {}
                  @gov_hash["LoanType/Term"]["Fixed"] = {}
                  @gov_hash["LoanType/Term"]["Fixed"]["16-29"] = {}
                  @gov_hash["LoanType/Term"]["Fixed"]["30"] = {}
                  cc = cc + 6
                  new_val = sheet_data.cell(r,cc)
                  @gov_hash["LoanType/Term"]["Fixed"]["16-29"] = new_val
                  @gov_hash["LoanType/Term"]["Fixed"]["30"] = new_val
                end
                if r >= 2261 && r <= 2266 && cc == 14
                  primary_key = get_value value
                  @gov_hash["FICO"][primary_key] = {}
                  cc = cc + 6
                  new_val = sheet_data.cell(r,cc)
                  @gov_hash["FICO"][primary_key] = new_val
                end
                if r == 2269 && cc == 14
                  @gov_hash["PropertyType"] = {}
                  @gov_hash["PropertyType"]["Manufactured Home"] = {}
                  cc = cc + 6
                  new_val = sheet_data.cell(r,cc)
                  @gov_hash["PropertyType"]["Manufactured Home"] = new_val
                end
                if r == 2270 && cc == 14
                  @gov_hash["PropertyType"]["3-4 Unit"] = {}
                  cc = cc + 6
                  new_val = sheet_data.cell(r,cc)
                  @gov_hash["PropertyType"]["3-4 Unit"] = new_val
                end
                if r == 2271 && cc == 14
                  @gov_hash["PropertyType"]["2nd Home"] = {}
                  @gov_hash["PropertyType"]["Investment Property"] = {}
                  cc = cc + 6
                  new_val = sheet_data.cell(r,cc)
                  @gov_hash["PropertyType"]["2nd Home"] = new_val
                  @gov_hash["PropertyType"]["Investment Property"] = new_val
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end
        adjustment = [@price_hash]
        make_adjust(adjustment,@gov_hash)
        #HECM / REVERSE MORTGAGE Not done

        #NON-QM: SIGMA SEASONED CREDIT EVENT, SIGMA RECENT CREDIT EVENT /program
        (2434..2477).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 7))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3*max_column + 2 # 2 / 5 / 8 / 11 / 14
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && (cc <= 8) && @title.class == String && @title != "N/A"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..50).each do |max_row|
                    @data = []
                    (0..2).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][30] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # #NON-QM: SIGMA SEASONED CREDIT EVENT, SIGMA RECENT CREDIT EVENT /adjustments
        # (range1_d..range2_d).each do |r|
        #   @ltv_data = sheet_data.row(2535)
        #   (0..sheet_data.last_column).each do |cc|
        #     begin
        #       value = sheet_data.cell(r,cc)
        #       if value == "ARM INFORMATION"
        #         primary_key = "LoanType/Term/LTV/FICO"
        #         first_row = 2436
        #         end_row = 2439
        #         first_column = 11
        #         last_column = 13
        #         ltv_row = 2435
        #         ltv_adjustment range1_d, range2_d, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "PROGRAM SPECIFIC PRICE ADJUSTMENTS"
        #         primary_key = "LoanType/Term/LTV/FICO"
        #         first_row = 2448
        #         end_row = 2464
        #         first_column = 11
        #         last_column = 20
        #         ltv_row = 2446
        #         ltv_adjustment range1_d, range2_d, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "SEASONED CREDIT EVENT"
        #         primary_key = "LoanType/Term/LTV/FICO"
        #         @spec_adjustment3[primary_key] = {}
        #       end

        #       if r >= 2537 && r <= 2544 && cc == 4
        #         ltv_key = get_value value
        #         @spec_adjustment3[primary_key][ltv_key] = {}
        #       end

        #       if r >= 2537 && r <= 2544 && cc >= 5 && cc <= 11
        #         c_val = get_value @ltv_data[cc-2]
        #         @spec_adjustment3[primary_key][ltv_key][c_val] = value
        #       end

        #       if value == " RECENT CREDIT EVENT"
        #         primary_key = "LoanType/Term/LTV/FICO"
        #         @spec_adjustment4[primary_key] = {}
        #       end

        #       if r >= 2537 && r <= 2544 && cc == 4
        #         ltv_key = get_value value
        #         @spec_adjustment4[primary_key][ltv_key] = {}
        #       end

        #       if r >= 2537 && r <= 2544 && cc >= 12 && cc <= 18
        #         c_val = get_value @ltv_data[cc-2]
        #         @spec_adjustment4[primary_key][ltv_key][c_val] = value
        #       end
        #     rescue Exception => e
        #       error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #       error_log.save
        #     end
        #   end
        # end

        # #NON-QM: SIGMA NO CREDIT EVENT PLUS /Program done
        # (2624..2675).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8
        #       begin
        #         @title = sheet_data.cell(r,cc)
        #         if @title.present? && (cc <= 8) && @title.class == String && @title != "N/A"
        #           @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #           program_property @program
        #           @programs_ids << @program.id
        #           @program.adjustments.destroy_all
        #           @block_hash = {}
        #           key = ''
        #           (1..50).each do |max_row|
        #             @data = []
        #             (0..2).each_with_index do |index, c_i|
        #               rrr = rr + max_row
        #               ccc = cc + c_i
        #               value = sheet_data.cell(rrr,ccc)
        #               if value.present?
        #                 if (c_i == 0)
        #                   key = value
        #                   @block_hash[key] = {}
        #                 else
        #                   @block_hash[key][30] = value
        #                 end
        #                 @data << value
        #               end
        #             end
        #             if @data.compact.reject { |c| c.blank? }.length == 0
        #               break # terminate the loop
        #             end
        #           end
        #           @program.update(base_rate: @block_hash)
        #         end
        #       rescue Exception => e
        #         error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #         error_log.save
        #       end
        #     end
        #   end
        # end

        # #NON-QM: SIGMA NO CREDIT EVENT PLUS /Adjustments 3 adjustments skip
        # (range1_e..range2_e).each do |r|
        #   @ltv_data = sheet_data.row(2730)
        #   (0..sheet_data.last_column).each do |cc|
        #     begin
        #       value = sheet_data.cell(r,cc)
        #       if value == "ARM INFORMATION"
        #         primary_key = "LoanType/Term/LTV/FICO"
        #         first_row = 2626
        #         end_row = 2629
        #         first_column = 11
        #         last_column = 13
        #         ltv_row = 2625
        #         ltv_adjustment range1_e, range2_e, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "PROGRAM SPECIFIC PRICE ADJUSTMENTS "
        #         primary_key = "LoanType/Term/LTV/FICO"
        #         first_row = 2638
        #         end_row = 2652
        #         first_column = 13
        #         last_column = 20
        #         ltv_row = 2636
        #         ltv_adjustment range1_e, range2_e, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "FULL DOCUMENTATION / ASSET UTILIZATION"
        #         primary_key = "LoanType/Term/LTV/FICO"
        #         @spec_adjustment5[primary_key] = {}
        #       end

        #       if r >= 2732 && r <= 2738 && cc == 2
        #         ltv_key = get_value value
        #         @spec_adjustment5[primary_key][ltv_key] = {}
        #       end

        #       if r >= 2732 && r <= 2738 && cc >= 3 && cc <= 11
        #         c_val = get_value @ltv_data[cc-2]
        #         @spec_adjustment5[primary_key][ltv_key][c_val] = value
        #       end

        #       if value == "BANK STATEMENT DOCUMENTION / EXPRESS DOCUMENTION"
        #         primary_key = "LoanType/Term/LTV/FICO"
        #         @spec_adjustment6[primary_key] = {}
        #       end

        #       if r >= 2732 && r <= 2738 && cc == 2
        #         ltv_key = get_value value
        #         @spec_adjustment6[primary_key][ltv_key] = {}
        #       end

        #       if r >= 2732 && r <= 2738 && cc >= 12 && cc <= 20
        #         c_val = get_value @ltv_data[cc-2]
        #         @spec_adjustment6[primary_key][ltv_key][c_val] = value
        #       end
        #     rescue Exception => e
        #       error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #       error_log.save
        #     end
        #   end
        # end

        # # NON-QM: R.E.A.L PRIME ADVANTAGE Programs done
        # (2791..2803).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8
        #       begin
        #         @title = sheet_data.cell(r,cc)
        #         if @title.present? && (cc <= 12) && @title.class == String #&& @title != "N/A"
        #           @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #           program_property @program
        #           @programs_ids << @program.id
        #           @program.adjustments.destroy_all
        #           @block_hash = {}
        #           key = ''
        #           (1..50).each do |max_row|
        #             @data = []
        #             (0..2).each_with_index do |index, c_i|
        #               rrr = rr + max_row
        #               ccc = cc + c_i
        #               value = sheet_data.cell(rrr,ccc)
        #               if value.present?
        #                 if (c_i == 0)
        #                   key = value
        #                   @block_hash[key] = {}
        #                 else
        #                   @block_hash[key][30] = value
        #                 end
        #                 @data << value
        #               end
        #             end
        #             if @data.compact.reject { |c| c.blank? }.length == 0
        #               break # terminate the loop
        #             end
        #           end
        #           @program.update(base_rate: @block_hash)
        #         end
        #       rescue Exception => e
        #         error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #         error_log.save
        #       end
        #     end
        #   end
        # end

        # #NON-QM: R.E.A.L PRIME ADVANTAGE /Adjustment Done
        # (range1_f..range2_f).each do |r|
        #   @ltv_data = sheet_data.row(2730)
        #   (0..sheet_data.last_column).each do |cc|
        #     begin
        #       value = sheet_data.cell(r,cc)
        #       if value == "ARM INFORMATION"
        #         primary_key = "LoanType/Term/LTV/FICO"
        #         first_row = 2799
        #         end_row = 2802
        #         first_column = 18
        #         last_column = 20
        #         ltv_row = 2798
        #         ltv_adjustment range1_f, range2_f, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
        #         primary_key = "LoanType / RateLock"
        #         @spec_adjustment7[primary_key] = {}
        #         c_val = sheet_data.cell(r,cc+4)
        #         @spec_adjustment7[primary_key][value] = c_val
        #       end

        #       # if r >= 2858 && r <= 2884 && cc == 16
        #       #   c_val = sheet_data.cell(r,cc+4)
        #       #   @spec_adjustment7[primary_key][value] = c_val
        #       # end

        #       if value == "FICO"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 2859
        #         end_row = 2865
        #         first_column = 2
        #         last_column = 12
        #         ltv_row = 2856
        #         ltv_adjustment range1_f, range2_f, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "LOAN AMOUNT"
        #         primary_key = "LoanType/LoanAmount/FICO"
        #         first_row = 2867
        #         end_row = 2874
        #         first_column = 2
        #         last_column = 12
        #         ltv_row = 2856
        #         ltv_adjustment range1_f, range2_f, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       # if value == "DOC TYPE" skip
        #       #   remaining
        #       # end

        #       if value == "45 Day Lock (Price Adjustment)"
        #         primary_key = "RateType/LoanType/RateLock"
        #         c_val = sheet_data.cell(2855,20)
        #         @day_adjustment[primary_key] = {}
        #         @day_adjustment[primary_key][value] = c_val
        #       end
        #     rescue Exception => e
        #       error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #       error_log.save
        #     end
        #   end
        # end

        # # NON-QM: R.E.A.L CREDIT ADVANTAGE - A /Program Done
        # (2936..2948).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8
        #       begin
        #         @title = sheet_data.cell(r,cc)
        #         if @title.present? && (cc <= 12) && @title.class == String #&& @title != "N/A"
        #           @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #           program_property @program
        #           @programs_ids << @program.id
        #           @program.adjustments.destroy_all
        #           @block_hash = {}
        #           key = ''
        #           (1..50).each do |max_row|
        #             @data = []
        #             (0..2).each_with_index do |index, c_i|
        #               rrr = rr + max_row
        #               ccc = cc + c_i
        #               value = sheet_data.cell(rrr,ccc)
        #               if value.present?
        #                 if (c_i == 0)
        #                   key = value
        #                   @block_hash[key] = {}
        #                 else
        #                   @block_hash[key][30] = value
        #                 end
        #                 @data << value
        #               end
        #             end
        #             if @data.compact.reject { |c| c.blank? }.length == 0
        #               break # terminate the loop
        #             end
        #           end
        #           @program.update(base_rate: @block_hash)
        #         end
        #       rescue Exception => e
        #         error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #         error_log.save
        #       end
        #     end
        #   end
        # end

        # # # NON-QM: R.E.A.L CREDIT ADVANTAGE - A //Adjustment
        # (range1_g..range2_g).each do |r|
        #   (0..sheet_data.last_column).each do |cc|
        #     begin
        #       value = sheet_data.cell(r,cc)
        #       if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
        #         primary_key = "LoanType / RateLock"
        #         @spec_adjustment12[primary_key] = {}
        #         c_val = sheet_data.cell(r,cc+4)
        #         @spec_adjustment12[primary_key][value] = c_val
        #       end

        #       # if r >= 3003 && r <= 3024 && cc == 16
        #       #   c_val = sheet_data.cell(r,cc+4)
        #       #   @spec_adjustment12[primary_key][value] = c_val
        #       # end

        #       if value == "45 Day Lock (Price Adjustment)"
        #         primary_key = "RateType/LoanType/RateLock"
        #         c_val = sheet_data.cell(3445,20)
        #         @day_adjustment6[primary_key] = {}
        #         @day_adjustment6[primary_key][value] = c_val
        #       end

        #       if value == "FICO"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 3004
        #         end_row = 3013
        #         first_column = 2
        #         last_column = 12
        #         ltv_row = 3001
        #         ltv_adjustment range1_g, range2_g, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "LOAN AMOUNT"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 3015
        #         end_row = 3023
        #         first_column = 2
        #         last_column = 12
        #         ltv_row = 3001
        #         ltv_adjustment range1_g, range2_g, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end
        #     rescue Exception => e
        #       error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #       error_log.save
        #     end
        #   end
        # end

        # # #NON-QM: R.E.A.L CREDIT ADVANTAGE - B, B-, C /Program
        # (3089..3101).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8
        #       begin
        #         @title = sheet_data.cell(r,cc)
        #         if @title.present? && (cc <= 12) && @title.class == String #&& @title != "N/A"
        #           @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #           program_property @program
        #           @programs_ids << @program.id
        #           @program.adjustments.destroy_all
        #           @block_hash = {}
        #           key = ''
        #           (1..50).each do |max_row|
        #             @data = []
        #             (0..2).each_with_index do |index, c_i|
        #               rrr = rr + max_row
        #               ccc = cc + c_i
        #               value = sheet_data.cell(rrr,ccc)
        #               if value.present?
        #                 if (c_i == 0)
        #                   key = value
        #                   @block_hash[key] = {}
        #                 else
        #                   @block_hash[key][30] = value
        #                 end
        #                 @data << value
        #               end
        #             end
        #             if @data.compact.reject { |c| c.blank? }.length == 0
        #               break # terminate the loop
        #             end
        #           end
        #           @program.update(base_rate: @block_hash)
        #         end
        #       rescue Exception => e
        #         error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #         error_log.save
        #       end
        #     end
        #   end
        # end

        # #NON-QM: R.E.A.L CREDIT ADVANTAGE - B, B-, C /Adjustment
        # (range1_h..range2_h).each do |r|
        #   (0..sheet_data.last_column).each do |cc|
        #     begin
        #       value = sheet_data.cell(r,cc)
        #       if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
        #         first_key = "LoanType / RateLock"
        #         @spec_adjustment8[first_key] = {}
        #       end

        #       if r >= 3158 && r <= 3181 && cc == 16
        #         c_val = sheet_data.cell(r,cc+4)
        #         @spec_adjustment8[first_key][value] = c_val
        #       end

        #       if value == "45 Day Lock (Price Adjustment)"
        #         primary_key = "RateType/LoanType/RateLock"
        #         c_val = sheet_data.cell(3155,20)
        #         @day_adjustment2[primary_key] = {}
        #         @day_adjustment2[primary_key][value] = c_val
        #       end

        #       if value == "FICO"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 3157
        #         end_row = 3167
        #         first_column = 2
        #         last_column = 11
        #         ltv_row = 3154
        #         ltv_adjustment range1_h, range2_h, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "LOAN AMOUNT"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 3169
        #         end_row = 3172
        #         first_column = 2
        #         last_column = 11
        #         ltv_row = 3154
        #         ltv_adjustment range1_h, range2_h, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       # if value == "DOC TYPE"
        #       #   remaining
        #       # end

        #       if value == "R.E.A.L Credit Advantage - B"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 3099
        #         end_row = 3102
        #         first_column = 14
        #         last_column = 16
        #         ltv_row = 3098
        #         ltv_adjustment range1_h, range2_h, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "R.E.A.L Credit Advantage - B- & C"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 3099
        #         end_row = 3102
        #         first_column = 18
        #         last_column = 20
        #         ltv_row = 3098
        #         ltv_adjustment range1_h, range2_h, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end
        #     rescue Exception => e
        #       error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #       error_log.save
        #     end
        #   end
        # end

        # #NON-QM: R.E.A.L INVESTOR INCOME - A /Program
        # (3237..3249).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8
        #       begin
        #         @title = sheet_data.cell(r,cc)
        #         if @title.present? && (cc <= 8) && @title.class == String #&& @title != "N/A"
        #           @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #           program_property @program
        #           @programs_ids << @program.id
        #           @program.adjustments.destroy_all
        #           @block_hash = {}
        #           key = ''
        #           (1..50).each do |max_row|
        #             @data = []
        #             (0..2).each_with_index do |index, c_i|
        #               rrr = rr + max_row
        #               ccc = cc + c_i
        #               value = sheet_data.cell(rrr,ccc)
        #               if value.present?
        #                 if (c_i == 0)
        #                   key = value
        #                   @block_hash[key] = {}
        #                 else
        #                   @block_hash[key][30] = value
        #                 end
        #                 @data << value
        #               end
        #             end
        #             if @data.compact.reject { |c| c.blank? }.length == 0
        #               break # terminate the loop
        #             end
        #           end
        #           @program.update(base_rate: @block_hash)
        #         end
        #       rescue Exception => e
        #         error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #         error_log.save
        #       end
        #     end
        #   end
        # end

        # #NON-QM: R.E.A.L INVESTOR INCOME - A /Adjustment
        # (range1_i..range2_i).each do |r|
        #   (0..sheet_data.last_column).each do |cc|
        #     begin
        #       value = sheet_data.cell(r,cc)
        #       if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
        #         first_key = "LoanType / RateLock"
        #         @spec_adjustment9[first_key] = {}
        #       end

        #       if r >= 3266 && r <= 3279 && cc == 16
        #         c_val = sheet_data.cell(r,cc+4)
        #         @spec_adjustment9[first_key][value] = c_val
        #       end

        #       if value == "45 Day Lock (Price Adjustment)"
        #         primary_key = "RateType/LoanType/RateLock"
        #         c_val = sheet_data.cell(3252,6)
        #         @day_adjustment3[primary_key] = {}
        #         @day_adjustment3[primary_key][value] = c_val
        #       end

        #       if value == "FICO"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 3257
        #         end_row = 3266
        #         first_column = 2
        #         last_column = 10
        #         ltv_row = 3254
        #         ltv_adjustment range1_i, range2_i, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "LOAN AMOUNT"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 3268
        #         end_row = 3275
        #         first_column = 2
        #         last_column = 10
        #         ltv_row = 3254
        #         ltv_adjustment range1_i, range2_i, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "DOC TYPE"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 3277
        #         end_row = 3280
        #         first_column = 2
        #         last_column = 10
        #         ltv_row = 3254
        #         ltv_adjustment range1_i, range2_i, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end
        #     rescue Exception => e
        #       error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #       error_log.save
        #     end
        #   end
        # end

        # #NON-QM: R.E.A.L INVESTOR INCOME - B, B- /Program
        # (3334..3346).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8
        #       begin
        #         @title = sheet_data.cell(r,cc)
        #         if @title.present? && (cc <= 12) && @title.class == String #&& @title != "N/A"
        #           @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #           program_property @program
        #           @programs_ids << @program.id
        #           @program.adjustments.destroy_all
        #           @block_hash = {}
        #           key = ''
        #           (1..50).each do |max_row|
        #             @data = []
        #             (0..2).each_with_index do |index, c_i|
        #               rrr = rr + max_row
        #               ccc = cc + c_i
        #               value = sheet_data.cell(rrr,ccc)
        #               if value.present?
        #                 if (c_i == 0)
        #                   key = value
        #                   @block_hash[key] = {}
        #                 else
        #                   @block_hash[key][30] = value
        #                 end
        #                 @data << value
        #               end
        #             end
        #             if @data.compact.reject { |c| c.blank? }.length == 0
        #               break # terminate the loop
        #             end
        #           end
        #           @program.update(base_rate: @block_hash)
        #         end
        #       rescue Exception => e
        #         error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #         error_log.save
        #       end
        #     end
        #   end
        # end

        # #NON-QM: R.E.A.L INVESTOR INCOME - B, B- /Adjustments
        # (range1_j..range2_j).each do |r|
        #   (0..sheet_data.last_column).each do |cc|
        #     begin
        #       value = sheet_data.cell(r,cc)

        #       if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
        #         first_key = "LoanType / RateLock"
        #         @spec_adjustment10[first_key] = {}
        #       end

        #       if r >= 3358 && r <= 3378 && cc == 16
        #         c_val = sheet_data.cell(r,cc+4)
        #         @spec_adjustment10[first_key][value] = c_val
        #       end

        #       if value == "45 Day Lock (Price Adjustment)"
        #         primary_key = "RateType/LoanType/RateLock"
        #         c_val = sheet_data.cell(3355,20)
        #         @day_adjustment4[primary_key] = {}
        #         @day_adjustment4[primary_key][value] = c_val
        #       end

        #       if value == "FICO"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 3351
        #         end_row = 3361
        #         first_column = 2
        #         last_column = 9
        #         ltv_row = 3348
        #         ltv_adjustment range1_j, range2_j, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end

        #       if value == "LOAN AMOUNT"
        #         primary_key = "LoanType/LTV/FICO"
        #         first_row = 3363
        #         end_row = 3365
        #         first_column = 2
        #         last_column = 10
        #         ltv_row = 3254
        #         ltv_adjustment range1_j, range2_j, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       end
        #     rescue Exception => e
        #       error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #       error_log.save
        #     end
        #     # PREPAYMENT PENALTY #adjustment remaining
        #   end
        # end

        # # # NON-QM: R.E.A.L DSC RATIO /Programs
        # (3433..3445).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8
        #       begin
        #         @title = sheet_data.cell(r,cc)
        #         if @title.present? && (cc <= 15) && @title.class == String #&& @title != "N/A"
        #           @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #           program_property @program
        #           @programs_ids << @program.id
        #           @program.adjustments.destroy_all
        #           @block_hash = {}
        #           key = ''
        #           (1..50).each do |max_row|
        #             @data = []
        #             (0..2).each_with_index do |index, c_i|
        #               rrr = rr + max_row
        #               ccc = cc + c_i
        #               value = sheet_data.cell(rrr,ccc)
        #               if value.present?
        #                 if (c_i == 0)
        #                   key = value
        #                   @block_hash[key] = {}
        #                 else
        #                   @block_hash[key][30] = value
        #                 end
        #                 @data << value
        #               end
        #             end
        #             if @data.compact.reject { |c| c.blank? }.length == 0
        #               break # terminate the loop
        #             end
        #           end
        #           @program.update(base_rate: @block_hash)
        #         end
        #       rescue Exception => e
        #         error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #         error_log.save
        #       end
        #     end
        #   end
        # end

        # # NON-QM: R.E.A.L DSC RATIO /Adjustment
        # (range1_k..range2_k).each do |r|
        #   (0..sheet_data.last_column).each do |cc|
        #     begin
        #       value = sheet_data.cell(r,cc)
        #       # if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
        #       #   primary_key = "LoanType / RateLock"
        #       #   @spec_adjustment11[primary_key] = {}
        #       # end

        #       # if r >= 3508 && r <= 3526 && cc == 16
        #       #   c_val = sheet_data.cell(r,cc+4)
        #       #   @spec_adjustment11[primary_key][value] = c_val
        #       # end

        #       # if value == "45 Day Lock (Price Adjustment)"
        #       #   primary_key = "RateType/LoanType/RateLock"
        #       #   c_val = sheet_data.cell(3445,20)
        #       #   @day_adjustment5[primary_key] = {}
        #       #   @day_adjustment5[primary_key][value] = c_val
        #       # end

        #       # if value == "FICO"
        #       #   primary_key = "LoanType/LTV/FICO"
        #       #   first_row = 3512
        #       #   end_row = 3519
        #       #   first_column = 2
        #       #   last_column = 9
        #       #   ltv_row = 3509
        #       #   ltv_adjustment range1_k, range2_k, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       # end

        #       # if value == "LOAN AMOUNT"
        #       #   primary_key = "LoanType/LTV/FICO"
        #       #   first_row = 3521
        #       #   end_row = 3527
        #       #   first_column = 2
        #       #   last_column = 10
        #       #   ltv_row = 3509
        #       #   ltv_adjustment range1_k, range2_k, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
        #       # end
        #     rescue Exception => e
        #       error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
        #       error_log.save
        #     end
        #   end
        # end
      end
      create_program_association_with_adjustment(sheet)
    end
    redirect_to programs_ob_sun_west_wholesale_demo5907_path(@sheet_obj)
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
          value1 = "0-"+value1.split("<=").last.tr('A-Za-z%$><= ','')
        elsif value1.include?(">") || value1.include?("+")
          value1 = value1.split(">").last.tr('A-Z=% ', '')+"-Inf"
        else
          value1 = value1.tr('A-Za-z% ','')
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
      if @program.program_name.include?("30 Year") || @program.program_name.include?("30Yr") || @program.program_name.include?("30 Yr") || @program.program_name.include?("30/25 Year") || @program.program_name.include?("30 YR")
        term = 30
      elsif @program.program_name.include?("20 Year") || @program.program_name.include?("20 YR")
        term = 20
      elsif @program.program_name.include?("15 Year") || @program.program_name.include?("15 YR")
        term = 15
      elsif @program.program_name.include?("10 Year") || @program.program_name.include?("10 YR")
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
        @program.loan_limit_type << "High-Balance"
      end
      @program.save
      @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline,loan_category: @sheet_name)
    end

    def ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
      @adjustment_hash = {}
      @adjustment_hash[primary_key] = {}
      @adjustment_hash[primary_key]["Conforming"] = {}
      @adjustment_hash[primary_key]["Conforming"]["15-Inf"] = {}
      ltv_key = ''
      cltv_key = ''
      (range1..range2).each do |r|
        row = sheet_data.row(r)
        @ltv_data = sheet_data.row(ltv_row)
        if row.compact.count >= 1
          (0..last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value.present?
                if r >= first_row && r <= end_row && cc == first_column
                  ltv_key = get_value value
                  @adjustment_hash[primary_key]["Conforming"]["15-Inf"][ltv_key] = {}
                end
                if r >= first_row && r <= end_row && cc > first_column && cc <= last_column
                  cltv_key = get_value @ltv_data[cc-2]
                  @adjustment_hash[primary_key]["Conforming"]["15-Inf"][ltv_key][cltv_key] = {}
                  @adjustment_hash[primary_key]["Conforming"]["15-Inf"][ltv_key][cltv_key] = value
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end
      end
      adjustment = [@adjustment_hash]
      make_adjust(adjustment,sheet)
    end

    def cash_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
      @adjustment_hash = {}
      @adjustment_hash[primary_key] = {}
      @adjustment_hash[primary_key]["Conforming"] = {}
      @adjustment_hash[primary_key]["Conforming"]["Cash Out"] = {}
      ltv_key = ''
      cltv_key = ''
      (range1..range2).each do |r|
        row = sheet_data.row(r)
        @ltv_data = sheet_data.row(ltv_row)
        if row.compact.count >= 1
          (0..last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value.present?
                if r >= first_row && r <= end_row && cc == first_column
                  ltv_key = get_value value
                  @adjustment_hash[primary_key]["Conforming"]["Cash Out"][ltv_key] = {}
                end
                if r >= first_row && r <= end_row && cc > first_column && cc <= last_column
                  cltv_key = get_value @ltv_data[cc-2]
                  @adjustment_hash[primary_key]["Conforming"]["Cash Out"][ltv_key][cltv_key] = {}
                  @adjustment_hash[primary_key]["Conforming"]["Cash Out"][ltv_key][cltv_key] = value
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end
      end
      adjustment = [@adjustment_hash]
      make_adjust(adjustment,sheet)
    end

    def lpmi_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
      @adjustment_hash = {}
      @adjustment_hash[primary_key] = {}
      @adjustment_hash[primary_key][true] = {}
      @adjustment_hash[primary_key][true]["Conforming"] = {}
      ltv_key = ''
      cltv_key = ''
      (range1..range2).each do |r|
        row = sheet_data.row(r)
        @ltv_data = sheet_data.row(ltv_row)
        if row.compact.count >= 1
          (0..last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value.present?
                if r >= first_row && r <= end_row && cc == first_column
                  ltv_key = get_value value
                  @adjustment_hash[primary_key][true]["Conforming"][ltv_key] = {}
                end
                if r >= first_row && r <= end_row && cc > first_column && cc <= last_column
                  cltv_key = get_value @ltv_data[cc-2]
                  @adjustment_hash[primary_key][true]["Conforming"][ltv_key][cltv_key] = {}
                  @adjustment_hash[primary_key][true]["Conforming"][ltv_key][cltv_key] = value
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end
      end
      adjustment = [@adjustment_hash]
      make_adjust(adjustment,sheet)
    end

    def sub_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, cltv_column, last_column, ltv_row, primary_key
      @adjustment_hash = {}
      @adjustment_hash[primary_key] = {}
      @adjustment_hash[primary_key]["Subordinate Financing"] = {}
      @adjustment_hash[primary_key]["Subordinate Financing"]["Conforming"] = {}
      ltv_key = ''
      cltv_key = ''
      new_key = ''
      (range1..range2).each do |r|
        row = sheet_data.row(r)
        @ltv_data = sheet_data.row(ltv_row)
        if row.compact.count >= 1
          (0..last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value.present?
                if r >= first_row && r <= end_row && cc == first_column
                  ltv_key = get_value value
                  @adjustment_hash[primary_key]["Subordinate Financing"]["Conforming"][ltv_key] = {}
                end
                if r >= first_row && r <= end_row && cc == cltv_column
                  new_key = get_value value
                  @adjustment_hash[primary_key]["Subordinate Financing"]["Conforming"][ltv_key][new_key] = {}
                end
                if r >= first_row && r <= end_row && cc > cltv_column && cc <= last_column
                  cltv_key = get_value @ltv_data[cc-2]
                  @adjustment_hash[primary_key]["Subordinate Financing"]["Conforming"][ltv_key][new_key][cltv_key] = {}
                  @adjustment_hash[primary_key]["Subordinate Financing"]["Conforming"][ltv_key][new_key][cltv_key] = value
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end
      end
      adjustment = [@adjustment_hash]
      make_adjust(adjustment,sheet)
    end

    def make_adjust(block_hash, sheet)
      block_hash.each do |hash|
        if hash.keys.count > 1
          hash.keys.each do |key|
            second_hash = {}
            second_hash[key] = hash[key]
            Adjustment.create(data: second_hash,loan_category: sheet)
          end
        else
          Adjustment.create(data: hash,loan_category: sheet)
        end
      end
    end

    def create_program_association_with_adjustment(sheet)
      adjustment_list = Adjustment.where(loan_category: sheet)
      program_list = Program.where(loan_category: sheet)

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
