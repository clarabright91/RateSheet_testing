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
        @cred_adjustment = {}
        @spec_adjustment = {}
        @spec_adjustment1 = {}
        @spec_adjustment2 = {}
        @spec_adjustment3 = {}
        @spec_adjustment4 = {}
        @spec_adjustment5 = {}
        @spec_adjustment6 = {}
        @spec_adjustment7 = {}
        @spec_adjustment8 = {}
        @spec_adjustment9 = {}
        @spec_adjustment10 = {}
        @spec_adjustment11 = {}
        @spec_adjustment12 = {}
        @home_adjustment = {}
        @sub_ord_hash = {}
        @sub_ord_hash1 = {}
        @caps_adjustment = {}
        @maximum_interest = {}
        @gov_adustment = {}
        @day_adjustment = {}
        @day_adjustment2 = {}
        @day_adjustment3 = {}
        @day_adjustment4 = {}
        @day_adjustment5 = {}
        @day_adjustment6 = {}
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
        range1_a = 782
        range2_a = 799
        range1_b = 1100
        range2_b = 1199
        range1_c = 1599
        range2_c = 1614
        range1_d = 2434
        range2_d = 2544
        range1_e = 2624
        range2_e = 2738
        range1_f = 2791
        range2_f = 2885
        range1_g = 2936
        range2_g = 3036
        range1_h = 3089
        range2_h = 3183
        range1_i = 3237
        range2_i = 3280
        range1_j = 3334
        range2_j = 3379
        range1_k = 3433
        range2_k = 3527

        # Agency Conforming Programs
        (156..320).each do |r|
          row = sheet_data.row(r)
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
                          @program.lock_period << 15*(c_i+1)
                          @program.save
                        end
                        @block_hash[main_key][key][15*(c_i+1)] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # PRICE ADJUSTMENTS: CONFORMING PROGRAMS //adjustment
        (range1..range2).each do |r|
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "LOAN TERM > 15 YEARS"
                primary_key = "LoanType/Term/LTV/FICO"
                first_row = 377
                end_row = 384
                last_column = 10
                first_column = 2
                ltv_row = 375
                ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "CASH OUT REFINANCE "
                primary_key = "LoanType/RefinanceOption/LTV"
                first_row = 389
                end_row = 396
                first_column = 2
                last_column = 6
                ltv_row = 387
                ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
              end

              if value == "ADDITIONAL LPMI ADJUSTMENTS"
                primary_key = "LPMI/RefinanceOption/FICO"
                first_row = 390
                end_row = 393
                first_column = 9
                last_column = 12
                ltv_row = 388
                ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
              end

              if value == "LPMI COVERAGE BASED ADJUSTMENTS"
                primary_key = "LPMI/RefinanceOption/FICO"
                first_row = 399
                end_row = 404
                first_column = 9
                last_column = 12
                ltv_row = 397
                ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
              end


              # SUBORDINATE FINANCING
              if value == "SUBORDINATE FINANCING"
                @ltv_data = sheet_data.row(399)
                first_key = "FinancingType/LTV/CLTV/FICO"
                second_key = "Subordinate Financing"
                @sub_ord_hash[first_key] = {}
                @sub_ord_hash[first_key][second_key] = {}
              end
              if r >= 400 && r <= 404 && cc == 2
                ltv_key = get_value value
                cltv_key = sheet_data.cell(r,cc+2)
                @sub_ord_hash[first_key][second_key][ltv_key] = {}
                @sub_ord_hash[first_key][second_key][ltv_key][cltv_key] = {}
              end
              if r >= 400 && r <= 404 && cc >= 6 && cc <= 7
                c_val = get_value @ltv_data[cc-2]
                @sub_ord_hash[first_key][second_key][ltv_key][cltv_key][c_val] = value
              end
              # SUBORDINATE FINANCING end

              # PROGRAM SPECIFIC ADJUSTMENTS
              if value == "PROGRAM SPECIFIC ADJUSTMENTS"
                primary_key = "LPMI/RefinanceOption/FICO"
                @spec_adjustment[primary_key] = {}
              end

              if r >= 375 && r <= 391 && cc == 15
                # if value.include?("Loan Amount")
                #   value = value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
                # elsif value.include?("Cashout")
                #   value = "Cashout/Fico/ltv"
                # elsif value.include?("Condo")
                #   value = "Condo"
                # else
                #   value = get_value value
                # end
                c_val = sheet_data.cell(r,cc+5)
                @spec_adjustment[primary_key][value] = c_val
              end

              if value == "FNMA HomeReady - Adjustment Caps"
                primary_key = "LoanType/Term/LTV/FICO"
                @home_adjustment[primary_key] = {}
              end

              if r >= 397 && r <= 398 && cc == 15
                c_val = sheet_data.cell(r,cc+5)
                @home_adjustment[primary_key] = {}
                @home_adjustment[primary_key][value] = c_val
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end

        # FHLMC HOME POSSIBLE / HOMEONE / SUPER CONFORMING Programs
        (708..760).each do |r|
          row = sheet_data.row(r)
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
                          @program.lock_period << 15*(c_i+1)
                          @program.save
                        end
                        @block_hash[main_key][key][15*(c_i+1)] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # PRICE ADJUSTMENTS: FHLMC HOME POSSIBLE / HOMEONE / SUPER CONFORMING //adjustment // 3 more adjustment remaining for this programs
        (range1_a..range2_a).each do |r|
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)

              if value == "LOAN TERM > 15 YEARS"
                primary_key = "LoanType/Term/LTV/FICO"
                first_row = 785
                end_row = 791
                first_column = 2
                last_column = 11
                ltv_row = 783
                ltv_adjustment range1_a, range2_a, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
              end

              if value == "PROGRAM SPECIFIC ADJUSTMENTS"
                primary_key = "LPMI/RefinanceOption/FICO"
                @spec_adjustment1[primary_key] = {}
              end

              if r >= 783 && r <= 791 && cc == 15
                c_val = sheet_data.cell(r,cc+5)
                @spec_adjustment1[primary_key][value] = c_val
              end

              # # SUBORDINATE FINANCING
              if value == "SUBORDINATE FINANCING  (Applicable to HomeOne)"
                @ltv_data = sheet_data.row(794)
                first_key = "FinancingType/LTV/CLTV/FICO"
                second_key = "Subordinate Financing"
                @sub_ord_hash1[first_key] = {}
                @sub_ord_hash1[first_key][second_key] = {}
              end
              if r >= 795 && r <= 799 && cc == 15
                ltv_key = get_value value
                cltv_key = sheet_data.cell(r,cc+2)
                @sub_ord_hash1[first_key][second_key][ltv_key] = {}
                @sub_ord_hash1[first_key][second_key][ltv_key][cltv_key] = {}
              end
              if r >= 795 && r <= 799 && cc >= 19 && cc <= 20
                c_val = get_value @ltv_data[cc-2]
                @sub_ord_hash1[first_key][second_key][ltv_key][cltv_key][c_val] = value
              end
              # # SUBORDINATE FINANCING end

              #ADJUSTMENT CAPS (Applicable to Home Possible products) Not completed
              if value == "ADJUSTMENT CAPS (Applicable to Home Possible products)"
                primary_key = "LoanType/Term/LTV/FICO"
                @caps_adjustment[primary_key] = {}
              end
              # ADJUSTMENT CAPS end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end

        #Non-Confirming: Sigma Programs
        (1101..1179).each do |r|
          row = sheet_data.row(r)
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
                          @program.lock_period << 15*(c_i+1)
                          @program.save
                        end
                        @block_hash[main_key][key][15*(c_i+1)] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #NON-CONFORMING: SIGMA QM PRIME JUMBO //adjustment
        (range1_b..range2_b).each do |r|
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)

              if value == "PRICE ADJUSTMENTS: <= $1,000,000"
                primary_key = "0 <= $1,000,000"
                first_row = 1185
                end_row = 1189
                first_column = 2
                last_column = 7
                ltv_row = 1183
                ltv_adjustment range1_b, range2_b, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
              end
              if value == "PRICE ADJUSTMENTS: > $1,000,000 & <= $1,500,000"
                primary_key = "$1,000,000"
                first_row = 1185
                end_row = 1189
                first_column = 2
                last_column = 12
                ltv_row = 1183
                ltv_adjustment range1_b, range2_b, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
              end

              if value == "PRICE ADJUSTMENTS: > $1,500,000 & <= $2,000,000"
                primary_key = "$1,500,000"
                first_row = 1195
                end_row = 1198
                first_column = 2
                last_column = 7
                ltv_row = 1193
                ltv_adjustment range1_b, range2_b, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
              end
              if value == "PRICE ADJUSTMENTS: > $2,000,000 & <= $2,500,000"
                primary_key = "$2,000,000"
                first_row = 1195
                end_row = 1198
                first_column = 2
                last_column = 12
                ltv_row = 1193
                ltv_adjustment range1_b, range2_b, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
              end


              if value == "PROGRAM SPECIFIC PRICE ADJUSTMENTS"
                primary_key = "$1,000,000"
                first_row = 1169
                end_row = 1182
                first_column = 14
                last_column = 20
                ltv_row = 1167
                ltv_adjustment range1_b, range2_b, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end

        #NON-CONFORMING: JW /Programs
        (1386..1547).each do |r|
          row = sheet_data.row(r)
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
                          @program.lock_period << 15*(c_i+1)
                          @program.save
                        end
                        @block_hash[main_key][key][15*(c_i+1)] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #NON-CONFORMING: JW /adjustment
        (range1_c..range2_c).each do |r|
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "ARM INFORMATION"  #Not done
                primary_key = "LoanType/Term/LTV/FICO"
                first_row = 377
                end_row = 384
                last_column = 10
                first_column = 2
                ltv_row = 375
                ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "PRICE ADJUSTMENTS"
                primary_key = "LoanType/Term/LTV/FICO"
                first_row = 1602
                end_row = 1607
                first_column = 7
                last_column = 11
                ltv_row = 1600
                ltv_adjustment range1_c, range2_c, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
              end

              if value == "STATE SPECEFIC PRICE ADJUSTMENTS"
                primary_key = "State"
                first_row = 1601
                end_row = 1610
                first_column = 13
                last_column = 20
                ltv_row = 1600
                ltv_adjustment range1_c, range2_c, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end

        #GOVERNMENT PROGRAMS /programs
        (2180..2203).each do |r|
          row = sheet_data.row(r)
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
                          @program.lock_period << 15*(c_i+1)
                          @program.save
                        end
                        @block_hash[main_key][key][15*(c_i+1)] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # #GOVERNMENT PROGRAMS /adjustments
        (2255..2278).each do |r|
          @ltv_data = sheet_data.row(2274)
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "PROGRAM SPECIFIC PRICE ADJUSTMENTS"
                primary_key = "LPMI/RefinanceOption/FICO"
                @spec_adjustment2[primary_key] = {}
              end

              if r >= 2256 && r <= 2271 && cc == 14
                c_val = sheet_data.cell(r,cc+6)
                @spec_adjustment2[primary_key][value] = c_val
              end

              if value == "Maximum Interest Rate allowed on USDA Product"
                primary_key = "MaximumInterest"
                c_val = sheet_data.cell(r,cc+2)
                @maximum_interest[primary_key] = {}
                @maximum_interest[primary_key][value] = c_val
              end

              if value == "Program"
                primary_key = "RateType/LoanType"
                @gov_adustment[primary_key] = {}
              end

              if r >= 2275 && r <= 2278 && cc == 16
                ltv_key = get_value value
                @gov_adustment[primary_key][ltv_key] = {}
              end

              if r >= 2275 && r <= 2278 && cc >= 19 && cc <= 20
                c_val = get_value @ltv_data[cc-2]
                @gov_adustment[primary_key][ltv_key][c_val] = value
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end

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
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #NON-QM: SIGMA SEASONED CREDIT EVENT, SIGMA RECENT CREDIT EVENT /adjustments
        (range1_d..range2_d).each do |r|
          @ltv_data = sheet_data.row(2535)
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "ARM INFORMATION"
                primary_key = "LoanType/Term/LTV/FICO"
                first_row = 2436
                end_row = 2439
                first_column = 11
                last_column = 13
                ltv_row = 2435
                ltv_adjustment range1_d, range2_d, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "PROGRAM SPECIFIC PRICE ADJUSTMENTS"
                primary_key = "LoanType/Term/LTV/FICO"
                first_row = 2448
                end_row = 2464
                first_column = 11
                last_column = 20
                ltv_row = 2446
                ltv_adjustment range1_d, range2_d, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "SEASONED CREDIT EVENT"
                primary_key = "LoanType/Term/LTV/FICO"
                @spec_adjustment3[primary_key] = {}
              end

              if r >= 2537 && r <= 2544 && cc == 4
                ltv_key = get_value value
                @spec_adjustment3[primary_key][ltv_key] = {}
              end

              if r >= 2537 && r <= 2544 && cc >= 5 && cc <= 11
                c_val = get_value @ltv_data[cc-2]
                @spec_adjustment3[primary_key][ltv_key][c_val] = value
              end

              if value == " RECENT CREDIT EVENT"
                primary_key = "LoanType/Term/LTV/FICO"
                @spec_adjustment4[primary_key] = {}
              end

              if r >= 2537 && r <= 2544 && cc == 4
                ltv_key = get_value value
                @spec_adjustment4[primary_key][ltv_key] = {}
              end

              if r >= 2537 && r <= 2544 && cc >= 12 && cc <= 18
                c_val = get_value @ltv_data[cc-2]
                @spec_adjustment4[primary_key][ltv_key][c_val] = value
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end

        #NON-QM: SIGMA NO CREDIT EVENT PLUS /Program done
        (2624..2675).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 7))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3*max_column + 2 # 2 / 5 / 8
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
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #NON-QM: SIGMA NO CREDIT EVENT PLUS /Adjustments 3 adjustments skip
        (range1_e..range2_e).each do |r|
          @ltv_data = sheet_data.row(2730)
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "ARM INFORMATION"
                primary_key = "LoanType/Term/LTV/FICO"
                first_row = 2626
                end_row = 2629
                first_column = 11
                last_column = 13
                ltv_row = 2625
                ltv_adjustment range1_e, range2_e, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "PROGRAM SPECIFIC PRICE ADJUSTMENTS "
                primary_key = "LoanType/Term/LTV/FICO"
                first_row = 2638
                end_row = 2652
                first_column = 13
                last_column = 20
                ltv_row = 2636
                ltv_adjustment range1_e, range2_e, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "FULL DOCUMENTATION / ASSET UTILIZATION"
                primary_key = "LoanType/Term/LTV/FICO"
                @spec_adjustment5[primary_key] = {}
              end

              if r >= 2732 && r <= 2738 && cc == 2
                ltv_key = get_value value
                @spec_adjustment5[primary_key][ltv_key] = {}
              end

              if r >= 2732 && r <= 2738 && cc >= 3 && cc <= 11
                c_val = get_value @ltv_data[cc-2]
                @spec_adjustment5[primary_key][ltv_key][c_val] = value
              end

              if value == "BANK STATEMENT DOCUMENTION / EXPRESS DOCUMENTION"
                primary_key = "LoanType/Term/LTV/FICO"
                @spec_adjustment6[primary_key] = {}
              end

              if r >= 2732 && r <= 2738 && cc == 2
                ltv_key = get_value value
                @spec_adjustment6[primary_key][ltv_key] = {}
              end

              if r >= 2732 && r <= 2738 && cc >= 12 && cc <= 20
                c_val = get_value @ltv_data[cc-2]
                @spec_adjustment6[primary_key][ltv_key][c_val] = value
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end

        # NON-QM: R.E.A.L PRIME ADVANTAGE Programs done
        (2791..2803).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 7))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3*max_column + 2 # 2 / 5 / 8
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && (cc <= 12) && @title.class == String #&& @title != "N/A"
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
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #NON-QM: R.E.A.L PRIME ADVANTAGE /Adjustment Done
        (range1_f..range2_f).each do |r|
          @ltv_data = sheet_data.row(2730)
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "ARM INFORMATION"
                primary_key = "LoanType/Term/LTV/FICO"
                first_row = 2799
                end_row = 2802
                first_column = 18
                last_column = 20
                ltv_row = 2798
                ltv_adjustment range1_f, range2_f, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
                primary_key = "LoanType / RateLock"
                @spec_adjustment7[primary_key] = {}
                c_val = sheet_data.cell(r,cc+4)
                @spec_adjustment7[primary_key][value] = c_val
              end

              # if r >= 2858 && r <= 2884 && cc == 16
              #   c_val = sheet_data.cell(r,cc+4)
              #   @spec_adjustment7[primary_key][value] = c_val
              # end

              if value == "FICO"
                primary_key = "LoanType/LTV/FICO"
                first_row = 2859
                end_row = 2865
                first_column = 2
                last_column = 12
                ltv_row = 2856
                ltv_adjustment range1_f, range2_f, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "LOAN AMOUNT"
                primary_key = "LoanType/LoanAmount/FICO"
                first_row = 2867
                end_row = 2874
                first_column = 2
                last_column = 12
                ltv_row = 2856
                ltv_adjustment range1_f, range2_f, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              # if value == "DOC TYPE" skip
              #   remaining
              # end

              if value == "45 Day Lock (Price Adjustment)"
                primary_key = "RateType/LoanType/RateLock"
                c_val = sheet_data.cell(2855,20)
                @day_adjustment[primary_key] = {}
                @day_adjustment[primary_key][value] = c_val
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end

        # NON-QM: R.E.A.L CREDIT ADVANTAGE - A /Program Done
        (2936..2948).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 7))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3*max_column + 2 # 2 / 5 / 8
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && (cc <= 12) && @title.class == String #&& @title != "N/A"
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
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # # NON-QM: R.E.A.L CREDIT ADVANTAGE - A //Adjustment
        (range1_g..range2_g).each do |r|
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
                primary_key = "LoanType / RateLock"
                @spec_adjustment12[primary_key] = {}
                c_val = sheet_data.cell(r,cc+4)
                @spec_adjustment12[primary_key][value] = c_val
              end

              # if r >= 3003 && r <= 3024 && cc == 16
              #   c_val = sheet_data.cell(r,cc+4)
              #   @spec_adjustment12[primary_key][value] = c_val
              # end

              if value == "45 Day Lock (Price Adjustment)"
                primary_key = "RateType/LoanType/RateLock"
                c_val = sheet_data.cell(3445,20)
                @day_adjustment6[primary_key] = {}
                @day_adjustment6[primary_key][value] = c_val
              end

              if value == "FICO"
                primary_key = "LoanType/LTV/FICO"
                first_row = 3004
                end_row = 3013
                first_column = 2
                last_column = 12
                ltv_row = 3001
                ltv_adjustment range1_g, range2_g, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "LOAN AMOUNT"
                primary_key = "LoanType/LTV/FICO"
                first_row = 3015
                end_row = 3023
                first_column = 2
                last_column = 12
                ltv_row = 3001
                ltv_adjustment range1_g, range2_g, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end

        # #NON-QM: R.E.A.L CREDIT ADVANTAGE - B, B-, C /Program
        (3089..3101).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 7))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3*max_column + 2 # 2 / 5 / 8
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && (cc <= 12) && @title.class == String #&& @title != "N/A"
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
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #NON-QM: R.E.A.L CREDIT ADVANTAGE - B, B-, C /Adjustment
        (range1_h..range2_h).each do |r|
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
                first_key = "LoanType / RateLock"
                @spec_adjustment8[first_key] = {}
              end

              if r >= 3158 && r <= 3181 && cc == 16
                c_val = sheet_data.cell(r,cc+4)
                @spec_adjustment8[first_key][value] = c_val
              end

              if value == "45 Day Lock (Price Adjustment)"
                primary_key = "RateType/LoanType/RateLock"
                c_val = sheet_data.cell(3155,20)
                @day_adjustment2[primary_key] = {}
                @day_adjustment2[primary_key][value] = c_val
              end

              if value == "FICO"
                primary_key = "LoanType/LTV/FICO"
                first_row = 3157
                end_row = 3167
                first_column = 2
                last_column = 11
                ltv_row = 3154
                ltv_adjustment range1_h, range2_h, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "LOAN AMOUNT"
                primary_key = "LoanType/LTV/FICO"
                first_row = 3169
                end_row = 3172
                first_column = 2
                last_column = 11
                ltv_row = 3154
                ltv_adjustment range1_h, range2_h, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              # if value == "DOC TYPE"
              #   remaining
              # end

              if value == "R.E.A.L Credit Advantage - B"
                primary_key = "LoanType/LTV/FICO"
                first_row = 3099
                end_row = 3102
                first_column = 14
                last_column = 16
                ltv_row = 3098
                ltv_adjustment range1_h, range2_h, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "R.E.A.L Credit Advantage - B- & C"
                primary_key = "LoanType/LTV/FICO"
                first_row = 3099
                end_row = 3102
                first_column = 18
                last_column = 20
                ltv_row = 3098
                ltv_adjustment range1_h, range2_h, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end

        #NON-QM: R.E.A.L INVESTOR INCOME - A /Program
        (3237..3249).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 7))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3*max_column + 2 # 2 / 5 / 8
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && (cc <= 8) && @title.class == String #&& @title != "N/A"
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
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #NON-QM: R.E.A.L INVESTOR INCOME - A /Adjustment
        (range1_i..range2_i).each do |r|
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
                first_key = "LoanType / RateLock"
                @spec_adjustment9[first_key] = {}
              end

              if r >= 3266 && r <= 3279 && cc == 16
                c_val = sheet_data.cell(r,cc+4)
                @spec_adjustment9[first_key][value] = c_val
              end

              if value == "45 Day Lock (Price Adjustment)"
                primary_key = "RateType/LoanType/RateLock"
                c_val = sheet_data.cell(3252,6)
                @day_adjustment3[primary_key] = {}
                @day_adjustment3[primary_key][value] = c_val
              end

              if value == "FICO"
                primary_key = "LoanType/LTV/FICO"
                first_row = 3257
                end_row = 3266
                first_column = 2
                last_column = 10
                ltv_row = 3254
                ltv_adjustment range1_i, range2_i, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "LOAN AMOUNT"
                primary_key = "LoanType/LTV/FICO"
                first_row = 3268
                end_row = 3275
                first_column = 2
                last_column = 10
                ltv_row = 3254
                ltv_adjustment range1_i, range2_i, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "DOC TYPE"
                primary_key = "LoanType/LTV/FICO"
                first_row = 3277
                end_row = 3280
                first_column = 2
                last_column = 10
                ltv_row = 3254
                ltv_adjustment range1_i, range2_i, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end

        #NON-QM: R.E.A.L INVESTOR INCOME - B, B- /Program
        (3334..3346).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 7))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3*max_column + 2 # 2 / 5 / 8
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && (cc <= 12) && @title.class == String #&& @title != "N/A"
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
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #NON-QM: R.E.A.L INVESTOR INCOME - B, B- /Adjustments
        (range1_j..range2_j).each do |r|
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)

              if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
                first_key = "LoanType / RateLock"
                @spec_adjustment10[first_key] = {}
              end

              if r >= 3358 && r <= 3378 && cc == 16
                c_val = sheet_data.cell(r,cc+4)
                @spec_adjustment10[first_key][value] = c_val
              end

              if value == "45 Day Lock (Price Adjustment)"
                primary_key = "RateType/LoanType/RateLock"
                c_val = sheet_data.cell(3355,20)
                @day_adjustment4[primary_key] = {}
                @day_adjustment4[primary_key][value] = c_val
              end

              if value == "FICO"
                primary_key = "LoanType/LTV/FICO"
                first_row = 3351
                end_row = 3361
                first_column = 2
                last_column = 9
                ltv_row = 3348
                ltv_adjustment range1_j, range2_j, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end

              if value == "LOAN AMOUNT"
                primary_key = "LoanType/LTV/FICO"
                first_row = 3363
                end_row = 3365
                first_column = 2
                last_column = 10
                ltv_row = 3254
                ltv_adjustment range1_j, range2_j, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
            # PREPAYMENT PENALTY #adjustment remaining
          end
        end

        # # NON-QM: R.E.A.L DSC RATIO /Programs
        (3433..3445).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 7))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3*max_column + 2 # 2 / 5 / 8
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && (cc <= 15) && @title.class == String #&& @title != "N/A"
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
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # NON-QM: R.E.A.L DSC RATIO /Adjustment
        (range1_k..range2_k).each do |r|
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              # if value == "PROGRAM SPECIFIC RATE ADJUSTMENTS"
              #   primary_key = "LoanType / RateLock"
              #   @spec_adjustment11[primary_key] = {}
              # end

              # if r >= 3508 && r <= 3526 && cc == 16
              #   c_val = sheet_data.cell(r,cc+4)
              #   @spec_adjustment11[primary_key][value] = c_val
              # end

              # if value == "45 Day Lock (Price Adjustment)"
              #   primary_key = "RateType/LoanType/RateLock"
              #   c_val = sheet_data.cell(3445,20)
              #   @day_adjustment5[primary_key] = {}
              #   @day_adjustment5[primary_key][value] = c_val
              # end

              # if value == "FICO"
              #   primary_key = "LoanType/LTV/FICO"
              #   first_row = 3512
              #   end_row = 3519
              #   first_column = 2
              #   last_column = 9
              #   ltv_row = 3509
              #   ltv_adjustment range1_k, range2_k, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              # end

              # if value == "LOAN AMOUNT"
              #   primary_key = "LoanType/LTV/FICO"
              #   first_row = 3521
              #   end_row = 3527
              #   first_column = 2
              #   last_column = 10
              #   ltv_row = 3509
              #   ltv_adjustment range1_k, range2_k, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, primary_key
              # end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
              error_log.save
            end
          end
        end
      end
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
        if value1.include?("FICO <") || value1.include?("FICO >=")
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
        @program.loan_limit_type << "High Balance"
      end
      @program.save
      @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline)
    end

    def ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row, primary_key
      @adjustment_hash = {}
      @adjustment_hash[primary_key] = {}
      # primary_key = ''
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
                  ltv_key = value
                  @adjustment_hash[primary_key][ltv_key] = {}
                end
                if r >= first_row && r <= end_row && cc > first_column && cc <= last_column
                  cltv_key = get_value @ltv_data[cc-2]
                  @adjustment_hash[primary_key][ltv_key][cltv_key] = {}
                  @adjustment_hash[primary_key][ltv_key][cltv_key] = value
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet, error_detail: e.message)
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
        Adjustment.create(data: hash,sheet_name: sheet)
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
