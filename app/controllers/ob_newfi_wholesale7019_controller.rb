class ObNewfiWholesale7019Controller < ApplicationController
  before_action :get_sheet, only: [:programs, :biscayne_delegated_jumbo, :sequoia_portfolio_plus_products, :sequoia_expanded_products, :sequoia_investor_pro, :fha_buydown_fixed_rate_products, :fha_fixed_arm_products, :fannie_mae_homeready_products, :fnma_buydown_products, :fnma_conventional_fixed_rate, :fnma_conventional_high_balance, :fnma_conventional_arm, :olympic_piggyback_fixed, :olympic_piggyback_high_balance, :olympic_piggyback_arm]
  before_action :get_program, only: [:single_program]

  def index
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "BISCAYNE DELEGATED JUMBO")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "Newfi Wholesale"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def biscayne_delegated_jumbo
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "BISCAYNE DELEGATED JUMBO")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @ltv_data = []
        @cltv_data = []
        @adjustment_hash = {}
        @purpose_adjustment = {}
        @highAdjustment = {}
        ltv_key = ''
        cltv_key = ''
        primary_key = ''

        #program
        (51..92).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count < 4)) || row.include?("7/1 LIBOR ARM BISCAYNE JUMBO") || row.include?("10/1 LIBOR ARM BISCAYNE JUMBO")
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              if r == 77
                cc = 15
              else
                cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              end

              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != "Rate"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property @program, sheet
                @programs_ids << @program.id
              

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''

                (1..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    begin
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        elsif (c_i == 1)
                          @block_hash[key][30] = value
                        elsif (c_i == 2)
                          @block_hash[key][45] = value
                        elsif (c_i == 3)
                          @block_hash[key][60] = value
                        else
                          @block_hash[key][15*c_i] = value
                        end
                        @data << value
                      end
                    rescue Exception => e
                      error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet)
                      error_log.save
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
        # adjustments
        (108..153).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(109)
          @cltv_data = sheet_data.row(136)
          (0..13).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Biscayne High Balance Price Adjustments"
                  primary_key = "HighBalance/FICO/CLTV"
                  @highAdjustment[primary_key] = {}
                end
                if r == 108 && value == "CLTV"
                  primary_key = "FICO/CLTV"
                  @adjustment_hash[primary_key] = {}
                end
                # CLTV
                if r >= 110 && r <= 115 && cc == 4
                  ltv_key = value
                  @adjustment_hash[primary_key][ltv_key] = {}
                end
                if r >= 110 && r <= 115 && cc >= 5 && cc <= 12
                  cltv_key = get_value @ltv_data[cc-3]
                  @adjustment_hash[primary_key][ltv_key][cltv_key] = {}
                  @adjustment_hash[primary_key][ltv_key][cltv_key] = value
                end
                if r >= 119 && r <= 125 && cc == 4
                  if value == "Cash-out Refinance"
                    primary_key = "RefinanceOption/LTV"
                  elsif value == "Purchase"
                    primary_key = "LoanPurpose/LTV"
                  else
                    primary_key = get_value value
                  end
                  @purpose_adjustment[primary_key] = {}
                end
                if r >= 119 && r <= 125 && cc >= 5 && cc <= 12
                  cltv_key = get_value @ltv_data[cc-3]
                  @purpose_adjustment[primary_key][cltv_key] = {}
                  @purpose_adjustment[primary_key][cltv_key] = value
                end
                # Biscayne High Balance Price Adjustments
                if r >= 137 && r <= 147 && r != 144 && r != 145 && cc == 4
                  if value == "Cash-Out Refi"
                    primary_key = "RefinanceOption/LTV"
                    @highAdjustment[primary_key] = {}
                    ltv_key = get_value value
                    @highAdjustment[primary_key][ltv_key] = {}
                  elsif value == "Purchase"
                    primary_key = "LoanPurpose/LTV"
                    @highAdjustment[primary_key] = {}
                    ltv_key = get_value value
                    @highAdjustment[primary_key][ltv_key] = {}
                  else
                    ltv_key = get_value value
                    @highAdjustment[primary_key][ltv_key] = {}
                  end
                end
                if r >= 137 && r <= 147 && r != 144 && r != 145 && cc >= 5 && cc <= 13
                  cltv_key = get_value @cltv_data[cc-3]
                  @highAdjustment[primary_key][ltv_key][cltv_key] = {}
                  @highAdjustment[primary_key][ltv_key][cltv_key] = value
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
              error_log.save
            end
          end
        end
        adjustment = [@adjustment_hash,@purpose_adjustment,@highAdjustment]
        make_adjust(adjustment,sheet)

        create_program_association_with_adjustment(sheet)
         # Update programs for Ob_SunWest sheet
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def sequoia_portfolio_plus_products
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "SEQUOIA PORTFOLIO PLUS PRODUCTS")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []

        #program
        (51..92).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id
                end

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
                error_log.save
              end
              # if @program.term.present?
              #   main_key = "Term/LoanType/InterestRate/LockPeriod"
              # else
              #   main_key = "InterestRate/LockPeriod"
              # end
              # @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  begin
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
                  rescue Exception => e
                    error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                    error_log.save
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
        # Adjustments
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def sequoia_expanded_products
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "SEQUOIA EXPANDED PRODUCTS")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []

        #program
        (51..92).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id
                end

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                error_log.save
              end
              # if @program.term.present?
              #   main_key = "Term/LoanType/InterestRate/LockPeriod"
              # else
              #   main_key = "InterestRate/LockPeriod"
              # end
              # @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  begin
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
                  rescue Exception => e
                    error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                    error_log.save
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

        # Adjustments
        (100..183).each do |r|
          row = sheet_data.row(r)
          if row.compact.count > 0
            (0..19).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?

              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def sequoia_investor_pro
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "SEQUOIA INVESTOR PRO")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @other_adjustment = {}
        primary_key = ''
        secondary_key = ''

        #program
        (51..92).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id
                end

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                error_log.save
              end
              # if @program.term.present?
              #   main_key = "Term/LoanType/InterestRate/LockPeriod"
              # else
              #   main_key = "InterestRate/LockPeriod"
              # end
              # @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  begin
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
                  rescue Exception => e
                    error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet)
                    error_log.save
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

        (77..113).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(81)
          if row.compact.count > 0
            (0..12).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "FICO x LTV"
                    primary_key = "LoanType/FICO/LTV"
                    @adjustment_hash[primary_key] = {}
                  end
                  # FICO x LTV
                  if r >= 82 && r <= 89 && cc == 5
                    secondary_key = value
                    @adjustment_hash[primary_key][secondary_key] = {}
                  end
                  if r >= 82 && r <= 89 && cc > 5 && cc <= 12
                    ltv_key = "#{(@ltv_data[cc-3]*100)}%"
                    @adjustment_hash[primary_key][secondary_key][ltv_key] = {}
                    @adjustment_hash[primary_key][secondary_key][ltv_key] = value
                  end
                  # Other
                  if r == 93 && cc == 5
                    primary_key = "LoanType"
                    @other_adjustment[primary_key] = {}
                  end
                  if r == 93 && cc > 5 && cc <= 12
                    ltv_key = "#{(@ltv_data[cc-3]*100)}%"
                    @other_adjustment[primary_key][ltv_key] = {}
                    @other_adjustment[primary_key][ltv_key] = value
                  end
                  if r >=94 && r <= 98 && cc == 4
                    primary_key = "LoanAmount"
                    @other_adjustment[primary_key] = {}
                  end
                  if r >=94 && r <= 98 && cc == 5
                    secondary_key = get_value value
                    @other_adjustment[primary_key][secondary_key] = {}
                  end
                  if r >= 94 && r <= 98 && cc > 5 && cc <= 12
                    ltv_key = "#{(@ltv_data[cc-3]*100)}%"
                    @other_adjustment[primary_key][secondary_key][ltv_key] = {}
                    @other_adjustment[primary_key][secondary_key][ltv_key] = value
                  end
                  if r >= 99 && r <= 105 && cc == 5
                    primary_key = value
                    @other_adjustment[primary_key] = {}
                  end
                  if r >= 99 && r <= 105 && cc > 5 && cc <= 12
                    ltv_key = "#{(@ltv_data[cc-3]*100)}%"
                    @other_adjustment[primary_key][ltv_key] = {}
                    @other_adjustment[primary_key][ltv_key] = value
                  end
                  if r >= 110 && r <= 113 && cc == 4
                    primary_key = value
                    @other_adjustment[primary_key] = {}
                  end
                  if r >= 110 && r <= 113 && cc == 5
                    @other_adjustment[primary_key] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@other_adjustment]
        make_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def fha_buydown_fixed_rate_products
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FHA BUYDOWN FIXED RATE PRODUCTS")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @other_adjustment = {}
        primary_key = ''
        secondary_key = ''
        ltv_key = ''

        #program
        (51..93).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4)) 
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property @program, sheet
                @programs_ids << @program.id
                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (1..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    begin
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
                    rescue Exception => e
                      error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet)
                      error_log.save
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

        # Adjustments
        (101..120).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(104)
          if row.compact.count >= 1
            (0..19).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "FICO - Loan Amount "
                    @adjustment_hash["FICO/LoanAmount"] = {}
                  end
                  # FICO - Loan Amount
                  if r >= 105 && r <= 110 && cc == 5
                    secondary_key = get_value value
                    @adjustment_hash["FICO/LoanAmount"][secondary_key] = {}
                  end
                  if r >= 105 && r <= 110 && cc > 5 && cc <= 8
                    ltv_key = get_value @ltv_data[cc-1]
                    if ltv_key == "0-110"
                      ltv_key = "0-110,000"
                    elsif ltv_key == "0-225"
                      ltv_key = "110,000-225,000"
                    elsif ltv_key == " 225-Inf"
                      ltv_key = "225-Inf"    
                    end
                    @adjustment_hash["FICO/LoanAmount"][secondary_key][ltv_key] = {}
                    @adjustment_hash["FICO/LoanAmount"][secondary_key][ltv_key] = value
                  end
                  # Other Adjustments
                  if r == 115 && cc == 14
                    @other_adjustment["MiscAdjuster"] = {}
                    @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = {}
                    cc = cc + 4
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = new_value
                  end
                  if r == 116 && cc == 14
                    @other_adjustment["LoanAmount"] = {}
                    @other_adjustment["LoanAmount"]["0-150,000"] = {}
                    cc = cc + 4
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["LoanAmount"]["0-150,000"] = new_value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@other_adjustment]
        make_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def fha_fixed_arm_products
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FHA FIXED ARM PRODUCTS")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []

        #program
        (51..115).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id
                end

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                error_log.save
              end
              # if @program.term.present?
              #   main_key = "Term/LoanType/InterestRate/LockPeriod"
              # else
              #   main_key = "InterestRate/LockPeriod"
              # end
              # @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  begin
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
                  rescue Exception => e
                    error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                    error_log.save
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
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def fannie_mae_homeready_products
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FANNIE MAE HOMEREADY PRODUCTS")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @price_adjustment = {}
        @family_adjustment = {}
        @high_adjustment = {}
        ltv_key = ''
        secondary_key = ''
        primary_key = ''
        ltv_key1 = ''
        secondary_key1 = ''
        primary_key1 = ''
        #program
        (51..92).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "5/1 CMT ARM 1/1/5 VA"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id

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
                  if @block_hash.keys.first.nil? || @block_hash.keys.first == "Rate"
                    @block_hash.shift
                  end
                  @program.update(base_rate: @block_hash)
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (101..128).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(104)
          @cltv_data = sheet_data.row(114)
          (0..16).each do |cc|
            value = sheet_data.cell(r,cc)
            begin
              if value.present?
                if value == " Price Adjustments"
                  primary_key = "FICO/LTV"
                  @price_adjustment[primary_key] = {}
                end
                if value == "Multi Family 2- 4 Unit LTV/FICO Adjusters"
                  @family_adjustment["PropertyType/FICO/LTV"] = {}
                  @family_adjustment["PropertyType/FICO/LTV"]["2-4 Unit"] = {}
                end
                if value == "Condo LTV/FICO Adjusters"
                  @family_adjustment["PropertyType/FICO/LTV"]["Condo"] = {}
                end
                if value == "HIGH BALANCE"
                  @high_adjustment["LoanSize/LoanType/LTV"] = {}
                  @high_adjustment["LoanSize/LoanType/LTV"]["High-Balance"] = {}
                  @high_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"] = {}
                  @high_adjustment["LoanType/RefinanceOption"] = {}
                  @high_adjustment["LoanType/RefinanceOption"]["Fixed"] = {}
                  @high_adjustment["LoanType/RefinanceOption"]["ARM"] = {}
                end
                if r >= 105 && r <= 112 && cc == 7
                  secondary_key = get_value value
                  @price_adjustment[primary_key][secondary_key] = {}
                end
                if r >= 105 && r <= 112 && cc > 7 && cc <= 112
                  ltv_key = get_value @ltv_data[cc-1]
                  if ltv_key.include?("%")
                    ltv_key = ltv_key.tr('% ','')
                  else
                    ltv_key
                  end
                  @price_adjustment[primary_key][secondary_key][ltv_key] = {}
                  @price_adjustment[primary_key][secondary_key][ltv_key] = value
                end
                # Multi Family 2- 4 Unit LTV/FICO Adjusters
                if r >= 115 && r <= 122 && cc == 6
                  ltv_key = get_value value
                  @family_adjustment["PropertyType/FICO/LTV"]["2-4 Unit"][ltv_key] = {}
                end
                if r >= 115 && r <= 122 && cc > 6 && cc <= 10
                  cltv_key = get_value @cltv_data[cc-1]
                  if cltv_key.include?("%")
                    cltv_key = cltv_key.tr('% ','')
                  else
                    cltv_key
                  end
                  @family_adjustment["PropertyType/FICO/LTV"]["2-4 Unit"][ltv_key][cltv_key] = {}
                  @family_adjustment["PropertyType/FICO/LTV"]["2-4 Unit"][ltv_key][cltv_key] = value
                end
                # Condo LTV/FICO Adjusters
                if r >= 115 && r <= 122 && cc == 12
                  ltv_key1 = get_value value
                  @family_adjustment["PropertyType/FICO/LTV"]["Condo"][ltv_key1] = {}
                end
                if r >= 115 && r <= 122 && cc > 12 && cc <= 16
                  cltv_key1 = get_value @cltv_data[cc-1]
                  if cltv_key1.include?("%")
                    cltv_key1 = cltv_key1.tr('% ','')
                  else
                    cltv_key1
                  end
                  @family_adjustment["PropertyType/FICO/LTV"]["Condo"][ltv_key1][cltv_key1] = {}
                  @family_adjustment["PropertyType/FICO/LTV"]["Condo"][ltv_key1][cltv_key1] = value
                end
                # HIGH BALANCE
                if r >= 124 && r <= 126 && cc == 9
                  if value.include?("LTV/CLTV >75% <=90%")
                    ltv_key = value.tr('A-Z/%>< ','').tr('=','-')
                  else
                    ltv_key = get_value value
                  end
                  @high_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"][ltv_key] = {}
                  cc = cc + 4
                  new_val = sheet_data.cell(r,cc)
                  @high_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"][ltv_key] = new_val
                end
                if r == 127 && cc == 9
                  ltv_key = "Rate and Term"
                  @high_adjustment["LoanType/RefinanceOption"]["Fixed"][ltv_key] = {}
                  @high_adjustment["LoanType/RefinanceOption"]["ARM"][ltv_key] = {}
                  cc = cc + 4
                  new_val = sheet_data.cell(r,cc)
                  @high_adjustment["LoanType/RefinanceOption"]["Fixed"][ltv_key] = new_val
                  @high_adjustment["LoanType/RefinanceOption"]["ARM"][ltv_key] = new_val
                end
                if r == 128 && cc == 9
                  ltv_key = "Cash Out"
                  @high_adjustment["LoanType/RefinanceOption"]["Fixed"][ltv_key] = {}
                  @high_adjustment["LoanType/RefinanceOption"]["ARM"][ltv_key] = {}
                  cc = cc + 4
                  new_val = sheet_data.cell(r,cc)
                  @high_adjustment["LoanType/RefinanceOption"]["Fixed"][ltv_key] = new_val
                  @high_adjustment["LoanType/RefinanceOption"]["ARM"][ltv_key] = new_val
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
              error_log.save
            end
          end
        end
        adjustment = [@price_adjustment,@family_adjustment,@high_adjustment]
        make_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def fnma_buydown_products
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FNMA BUYDOWN PRODUCTS")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        range1 = 101
        range2 = 131
        @other_adjustment = {}
        @secondary_hash = {}
        primary_key = ''
        ltv_key = ''
        cltv_key = ''
        ltv_data = ''

        #program
        (51..93).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id
                end

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                error_log.save
              end
              # if @program.term.present?
              #   main_key = "Term/LoanType/InterestRate/LockPeriod"
              # else
              #   main_key = "InterestRate/LockPeriod"
              # end
              # @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  begin
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
                  rescue Exception => e
                    error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet)
                    error_log.save
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

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(117)
          (0..sheet_data.last_column).each do |cc|
            value = sheet_data.cell(r,cc)
            begin
              if value == "LTV / FICO (Terms > 15 years only)"
                first_row = 105
                end_row = 108
                last_column = 13
                first_column = 5
                ltv_row = 104
                num = 3
                ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row,num
              end
              # Loans With Secondary Financing
              if value == "Loans With Secondary Financing"
                primary_key = "LTV/CLTV/FICO"
                @secondary_hash[primary_key] = {}
              end
              if r >= 118 && r <= 122 && cc == 4
                ltv_key = get_value value
                @secondary_hash[primary_key][ltv_key] = {}
              end
              if r >= 118 && r <= 122 && cc == 5
                cltv_key = get_value value
                @secondary_hash[primary_key][ltv_key][cltv_key] = {}
              end
              if r >= 118 && r <= 122 && cc > 5 && cc <= 7
                ltv_data = get_value @ltv_data[cc-3]
                ltv_data = ltv_data.tr(')( ','')
                @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = {}
                @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = value
              end
              # Other Adjustments
              if r == 108 && cc == 15
                @other_adjustment["PropertyType"] = {}
                @other_adjustment["PropertyType"]["2-4 Unit"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType"]["2-4 Unit"] = new_value
              end
              if r == 109 && cc == 15
                @other_adjustment["PropertyType/Term/LTV"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = new_value
              end
              if r == 110 && cc == 15
                @other_adjustment["LoanSize/RefinanceOption"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = new_value
              end
              if r == 111 && cc == 15
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Rate and Term"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Rate and Term"] = new_value
              end
              if r == 130 && cc == 14
                @other_adjustment["MiscAdjuster"] = {}
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = new_value
              end
              if r == 131 && cc == 14
                @other_adjustment["LoanAmount"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["0-150,000"] = new_value
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
              error_log.save
            end
          end
        end
        adjustment = [@secondary_hash,@other_adjustment]
        make_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def fnma_conventional_fixed_rate
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FNMA Conventional Fixed Rate")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        range1 = 137
        range2 = 178
        @cashout_hash = {}
        @lpmi_hash = {}
        @other_adjustment = {}
        @secondary_hash = {}
        primary_key = ''
        ltv_key = ''
        cltv_key = ''
        ltv_data = ''

        #program
        (51..93).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id
                end

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                # if @program.term.present?
                #   main_key = "Term/LoanType/InterestRate/LockPeriod"
                # else
                #   main_key = "InterestRate/LockPeriod"
                # end
                # @block_hash[main_key] = {}
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(153)
          @lpmi_data = sheet_data.row(166)
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "LTV / FICO (Terms > 15 years only)"
                first_row = 141
                end_row = 147
                last_column = 12
                first_column = 4
                ltv_row = 140
                num = 1
                ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row,num
              end
              if value == "Cash Out Refinance"
                @cashout_hash["RefinanceOption/FICO/LTV"] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
              end
              if value == "LPMI Single Premium Rate Card"
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["20-Inf"] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["0-20"] = {}
              end
              # Cash Out Refinance
              if r >= 154 && r <= 160 && cc == 4
                primary_key = get_value value
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key] = {}
              end
              if r >= 154 && r <= 160 && cc >= 5 && cc <= 8
                ltv_key = get_value @ltv_data[cc-1]
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = value
              end
              # LPMI Single Premium Rate Card
              if r >= 167 && r <= 170 && cc == 2
                if value.include?("below")
                  primary_key = "0-"+value.tr('a-z% ','')
                else
                  primary_key = value.sub('to','-').tr('% ','')
                end
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["20-Inf"][primary_key] = {}
              end
              if r >= 167 && r <= 170 && cc >= 5 && cc <= 11
                ltv_key = get_value @lpmi_data[cc-1]
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["20-Inf"][primary_key][ltv_key] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["20-Inf"][primary_key][ltv_key] = value
              end
              if r >= 173 && r <= 176 && cc == 2
                if value.include?("below")
                  primary_key = "0-"+value.tr('a-z% ','')
                else
                  primary_key = value.sub('to','-').tr('% ','')
                end
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["0-20"][primary_key] = {}
              end
              if r >= 173 && r <= 176 && cc >= 5 && cc <= 11
                ltv_key = get_value @lpmi_data[cc-1]
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["0-20"][primary_key][ltv_key] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["0-20"][primary_key][ltv_key] = value
              end
              if r == 178 && cc == 2
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"]["2nd Home"] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"]["2nd Home"]["Rate and Term"] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"]["2nd Home"]["Rate and Term"]["0-20"] = {}
              end
              if r == 178 && cc >= 5 && cc <= 11
                ltv_key = get_value @lpmi_data[cc-1]
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"]["2nd Home"]["Rate and Term"]["0-20"][ltv_key] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"]["2nd Home"]["Rate and Term"]["0-20"][ltv_key] = value
              end
              # Other Adjustments
              if r == 141 && cc == 14
                @other_adjustment["PropertyType/LTV"] = {}
                @other_adjustment["PropertyType/LTV"]["Investment Property"] = {}
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["0-75"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["0-75"] = {}
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["0-75"] = new_value
              end
              if r == 142 && cc == 14
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["75-80"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["75-80"] = {}
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["75-80"] = new_value
              end
              if r == 143 && cc == 14
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["80-Inf"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["80-Inf"] = {}
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["80-Inf"] = new_value
              end
              if r == 144 && cc == 14
                @other_adjustment["PropertyType"] = {}
                @other_adjustment["PropertyType"]["2-4 Unit"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType"]["2-4 Unit"] = new_value
              end
              if r == 145 && cc == 14
                @other_adjustment["PropertyType/LTV/Term"] = {}
                @other_adjustment["PropertyType/LTV/Term"]["Condo"] = {}
                @other_adjustment["PropertyType/LTV/Term"]["Condo"]["75-Inf"] = {}
                @other_adjustment["PropertyType/LTV/Term"]["Condo"]["75-Inf"]["15-Inf"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV/Term"]["Condo"]["75-Inf"]["15-Inf"] = new_value
              end
              if r == 146 && cc == 14
                @other_adjustment["LoanSize/RefinanceOption"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = new_value
              end
              if r == 147 && cc == 14
                @other_adjustment["LoanSize/LoanType/RefinanceOption"] = {}
                @other_adjustment["LoanSize/LoanType/RefinanceOption"]["High-Balance"] = {}
                @other_adjustment["LoanSize/LoanType/RefinanceOption"]["High-Balance"]["ARM"] = {}
                @other_adjustment["LoanSize/LoanType/RefinanceOption"]["High-Balance"]["ARM"]["Rate and Term"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/LoanType/RefinanceOption"]["High-Balance"]["ARM"]["Rate and Term"] = new_value
              end
              if r == 166 && cc == 13
                @other_adjustment["MiscAdjuster"] = {}
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = new_value
              end
              if r == 167 && cc == 13
                @other_adjustment["LoanAmount"] = {}
                @other_adjustment["LoanAmount"]["150,000"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["150,000"] = new_value
              end
              if r == 168 && cc == 13
                @other_adjustment["LoanAmount"]["250,000-Inf"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["250,000-Inf"] = new_value
              end
              if r == 169 && cc == 13
                @other_adjustment["LoanAmount"]["200,000-250,000"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["200,000-250,000"] = new_value
              end
              if r == 170 && cc == 13
                primary_key = "FICO"
                secondary_key =  "0-680"
                if @other_adjustment[primary_key] = {}
                  cc = cc + 4
                  new_value = sheet_data.cell(r,cc)
                  @other_adjustment[primary_key][secondary_key] = new_value
                end
              end
              if r == 171 && cc == 13
                @other_adjustment["LoanAmount/State"] = {}
                @other_adjustment["LoanAmount/State"]["275,000-Inf"] = {}
                @other_adjustment["LoanAmount/State"]["275,000-Inf"]["CA"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount/State"]["275,000-Inf"]["CA"] = new_value
              end
              if r == 172 && cc == 13
                @other_adjustment["LoanAmount/State"]["200,000-275,000"] = {}
                @other_adjustment["LoanAmount/State"]["200,000-275,000"]["CA"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount/State"]["200,000-275,000"]["CA"] = new_value
              end
              # Loans With Secondary Financing
              if value == "Loans With Secondary Financing"
                @secondary_hash["LTV/CLTV/FICO"] = {}
              end
              if r >= 154 && r <= 158 && cc == 12
                ltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key] = {}
              end
              if r >= 154 && r <= 158 && cc == 13
                cltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key] = {}
              end
              if r >= 154 && r <= 158 && cc > 13 && cc <= 15
                ltv_data = get_value @ltv_data[cc-1]
                ltv_data = ltv_data.tr('() ','')
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = {}
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = value
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
              error_log.save
            end
          end
        end
        adjustment = [@cashout_hash,@lpmi_hash,@secondary_hash,@other_adjustment]
        make_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def fnma_conventional_high_balance
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FNMA Conventional High Balance")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        range1 = 137
        range2 = 184
        @cashout_hash = {}
        @lpmi_hash = {}
        @other_adjustment = {}
        @secondary_hash = {}
        primary_key = ''
        ltv_key = ''
        cltv_key = ''
        ltv_data = ''

        #program
        (51..93).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id
                end

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                error_log.save
              end

              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  begin
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
                  rescue Exception => e
                    error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                    error_log.save
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

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(153)
          @lpmi_data = sheet_data.row(172)
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "LTV / FICO (Terms > 15 years only)"
                first_row = 141
                end_row = 147
                last_column = 12
                first_column = 4
                ltv_row = 140
                num = 1
                ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row,num
              end
              # Cash Out Refinance
              if value == "Cash Out Refinance"
                @cashout_hash["RefinanceOption/FICO/LTV"] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
              end
              # LPMI Single Premium Rate Card
              if value == "LPMI Single Premium Rate Card"
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["20-Inf"] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["0-20"] = {}
              end
              # Cash Out Refinance
              if r >= 154 && r <= 160 && cc == 4
                primary_key = get_value value
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key] = {}
              end
              if r >= 154 && r <= 160 && cc >= 5 && cc <= 8
                ltv_key = get_value @ltv_data[cc-1]
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = value
              end
              # LPMI Single Premium Rate Card
              if r >= 173 && r <= 176 && cc == 2
                if value.include?("below")
                  primary_key = "0-"+value.tr('a-z% ','')
                else
                  primary_key = value.sub('to','-').tr('% ','')
                end
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["20-Inf"][primary_key] = {}
              end
              if r >= 173 && r <= 176 && cc >= 5 && cc <= 11
                ltv_key = get_value @lpmi_data[cc-1]
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["20-Inf"][primary_key][ltv_key] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["20-Inf"][primary_key][ltv_key] = value
              end
              if r >= 179 && r <= 182 && cc == 2
                if value.include?("below")
                  primary_key = "0-"+value.tr('a-z% ','')
                else
                  primary_key = value.sub('to','-').tr('% ','')
                end
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["0-20"][primary_key] = {}
              end
              if r >= 179 && r <= 182 && cc >= 5 && cc <= 11
                ltv_key = get_value @lpmi_data[cc-1]
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["0-20"][primary_key][ltv_key] = {}
                @lpmi_hash["LPMI/LoanType/RefinanceOption/Term/FICO/LTV"][true]["Fixed"]["Rate and Term"]["0-20"][primary_key][ltv_key] = value
              end
              if r == 184 && cc == 2
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"]["2nd Home"] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"]["2nd Home"]["Rate and Term"] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"]["2nd Home"]["Rate and Term"]["0-20"] = {}
              end
              if r == 184 && cc >= 5 && cc <= 11
                ltv_key = get_value @lpmi_data[cc-1]
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"]["2nd Home"]["Rate and Term"]["0-20"][ltv_key] = {}
                @lpmi_hash["LPMI/LoanType/PropertyType/RefinanceOption/Term/LTV"][true]["Fixed"]["2nd Home"]["Rate and Term"]["0-20"][ltv_key] = value
              end
              # Other Adjustments
              if r == 141 && cc == 14
                @other_adjustment["PropertyType/LTV"] = {}
                @other_adjustment["PropertyType/LTV"]["Investment Property"] = {}
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["0-75"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["0-75"] = new_value
              end

              if r == 142 && cc == 14
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["75-80"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["75-80"] = new_value
              end
              if r == 143 && cc == 14
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["80-Inf"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["80-Inf"] = new_value
              end
              if r == 144 && cc == 14
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["2-4 Unit"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["2-4 Unit"] = new_value
              end
              if r == 145 && cc == 14
                @other_adjustment["PropertyType/LTV/Term"] = {}
                @other_adjustment["PropertyType/LTV/Term"]["Condo"] = {}
                @other_adjustment["PropertyType/LTV/Term"]["Condo"]["75-Inf"] = {}
                @other_adjustment["PropertyType/LTV/Term"]["Condo"]["75-Inf"]["15-Inf"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV/Term"]["Condo"]["75-Inf"]["15-Inf"] = new_value
              end
              if r == 146 && cc == 14
                @other_adjustment["LoanSize/RefinanceOption"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = new_value
              end
              if r == 147 && cc == 14
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Rate and Term"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Rate and Term"] = new_value
              end
              if r == 166 && cc == 14
                @other_adjustment["MiscAdjuster"] = {}
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = new_value
              end
              if r == 167 && cc == 14
                @other_adjustment["LoanAmount"] = {}
                @other_adjustment["LoanAmount"]["0-150,000"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["0-150,000"] = new_value
              end
              if r == 168 && cc == 14
                @other_adjustment["LoanAmount"]["300,000-Inf"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["300,000-Inf"] = new_value
              end
              if r == 169 && cc == 14
                @other_adjustment["LoanAmount"]["200,000-300,000"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["200,000-300,000"] = new_value
              end
              if r == 170 && cc == 14
                primary_key = "FICO"
                secondary_key =  "640-679"
                if @other_adjustment[primary_key] = {}
                  cc = cc + 1
                  new_value = sheet_data.cell(r,cc)
                  @other_adjustment[primary_key][secondary_key] = new_value
                end
              end
              # Loans With Secondary Financing
              if value == "Loans With Secondary Financing"
                @secondary_hash["LTV/CLTV/FICO"] = {}
              end
              if r >= 154 && r <= 158 && cc == 12
                ltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key] = {}
              end
              if r >= 154 && r <= 158 && cc == 13
                cltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key] = {}
              end
              if r >= 154 && r <= 158 && cc > 13 && cc <= 15
                ltv_data = get_value @ltv_data[cc-1]
                ltv_data = ltv_data.tr('() ','')
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = {}
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = value
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
              error_log.save
            end
          end
        end
        adjustment = [@cashout_hash,@secondary_hash,@other_adjustment,@lpmi_hash]
        make_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def fnma_conventional_arm
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FNMA Conventional Arm")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        range1 = 137
        range2 = 171
        @cashout_hash = {}
        @other_adjustment = {}
        @secondary_hash = {}
        primary_key = ''
        ltv_key = ''
        cltv_key = ''
        ltv_data = ''

        #program
        (51..93).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              @title = sheet_data.cell(r,cc)
              begin
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id
                end

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                lock_hash = {}
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
                error_log.save
              end

              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  begin
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
                  rescue Exception => e
                    error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet)
                    error_log.save
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

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(154)
          (0..sheet_data.last_column).each do |cc|
            value = sheet_data.cell(r,cc)
            begin
              if value == "LTV / FICO (Terms > 15 years only)"
                first_row = 141
                end_row = 147
                last_column = 12
                first_column = 4
                ltv_row = 140
                num = 1
                ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row,num
              end
              # Cash out Refinance
              if value == "Cash out Refinance"
                @cashout_hash["RefinanceOption/FICO/LTV"] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
              end
              # Cash Out Refinance
              if r >= 155 && r <= 161 && cc == 4
                primary_key = get_value value
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key] = {}
              end
              if r >= 155 && r <= 161 && cc >= 5 && cc <= 8
                ltv_key = get_value @ltv_data[cc-1]
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = value
              end
              # Other Adjustments
              if r == 141 && cc == 14
                @other_adjustment["PropertyType/LTV"] = {}
                @other_adjustment["PropertyType/LTV"]["Investment Property"] = {}
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["0-75"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["0-75"] = new_value
              end
              if r == 142 && cc == 14
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["75-80"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["75-80"] = new_value
              end
              if r == 143 && cc == 14
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["80-Inf"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/LTV"]["Investment Property"]["80-Inf"] = new_value
              end
              if r == 144 && cc == 14
                @other_adjustment["PropertyType"] = {}
                @other_adjustment["PropertyType"]["2-4 Unit"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType"]["2-4 Unit"] = new_value
              end
              if r == 145 && cc == 14
                @other_adjustment["PropertyType/Term/LTV"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = new_value
              end
              if r == 146 && cc == 14
                @other_adjustment["LoanSize/RefinanceOption"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = new_value
              end
              if r == 147 && cc == 14
                @other_adjustment["LoanSize/LoanType/LTV"] = {}
                @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"] = {}
                @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"] = {}
                @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["0-75"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["0-75"] = new_value
              end
              if r == 148 && cc == 14
                @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["75-90"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["75-90"] = new_value
              end
              if r == 149 && cc == 14
                @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["90-95"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["90-95"] = new_value
              end
              if r == 150 && cc == 14
                @other_adjustment["LoanSize/RefinanceOption"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Rate and Term"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Rate and Term"] = new_value
              end
              if r == 165 && cc == 12
                @other_adjustment["MiscAdjuster"] = {}
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = new_value
              end
              if r == 166 && cc == 12
                @other_adjustment["LoanAmount"] = {}
                @other_adjustment["LoanAmount"]["0-150,000"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["0-150,000"] = new_value
              end
              if r == 167 && cc == 12
                @other_adjustment["LoanAmount"]["250,000-Inf"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["250,000-Inf"] = new_value
              end
              if r == 168 && cc == 12
                @other_adjustment["LoanAmount"]["200,000-250,000"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["200,000-250,000"] = new_value
              end
              if r == 169 && cc == 12
                primary_key = "FICO"
                secondary_key =  "0-680"
                if @other_adjustment[primary_key] = {}
                  cc = cc + 4
                  new_value = sheet_data.cell(r,cc)
                  @other_adjustment[primary_key][secondary_key] = new_value
                end
              end
              if r == 170 && cc == 12
                @other_adjustment["LoanAmount/State"] = {}
                @other_adjustment["LoanAmount/State"]["275,000-Inf"] = {}
                @other_adjustment["LoanAmount/State"]["275,000-Inf"]["CA"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount/State"]["275,000-Inf"]["CA"] = new_value
              end
              if r == 171 && cc == 12
                @other_adjustment["LoanAmount/State"]["200,000-275,000"] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount/State"]["200,000-275,000"] = new_value
              end

              # Loans With Secondary Financing
              if value == "Loan With Secondary Financing"
                @secondary_hash["LTV/CLTV/FICO"] = {}
              end
              if r >= 155 && r <= 159 && cc == 12
                ltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key] = {}
              end
              if r >= 155 && r <= 159 && cc == 13
                cltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key] = {}
              end
              if r >= 155 && r <= 159 && cc > 13 && cc <= 15
                ltv_data = get_value @ltv_data[cc-1]
                ltv_data = ltv_data.tr('() ','')
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = {}
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = value
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
              error_log.save
            end
          end
        end
        adjustment = [@secondary_hash,@cashout_hash,@other_adjustment]
        make_adjust(adjustment,sheet)
        # create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def olympic_piggyback_fixed
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Olympic PiggyBack Fixed")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        range1 = 137
        range2 = 174
        @cashout_hash = {}
        @other_adjustment = {}
        @secondary_hash = {}
        @property_hash = {}
        primary_key = ''
        ltv_key = ''
        cltv_key = ''
        ltv_data = ''

        #program
        (51..93).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id
                end

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                error_log.save
              end

              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  begin
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
                  rescue Exception => e
                    error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet)
                    error_log.save
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

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(154)
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "LTV / FICO (Terms > 15 years only)"
                first_row = 141
                end_row = 144
                last_column = 8
                first_column = 4
                ltv_row = 140
                num = 1
                ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row,num
              end
              if value == "Cash Out Refinance"
                @cashout_hash["RefinanceOption/FICO/LTV"] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
                @property_hash["LoanType/CLTV"] = {}
                @property_hash["LoanType/CLTV"]["Fixed"] = {}
              end
              # Cash Out Refinance
              if r >= 155 && r <= 158 && cc == 4
                primary_key = get_value value
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key] = {}
              end
              if r >= 155 && r <= 158 && cc >= 5 && cc <= 8
                ltv_key = get_value @ltv_data[cc-1]
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = value
              end
              # OLYMPIC FIXED 2ND MORTGAGE
              if r == 168 && cc == 4
                @property_hash["LoanType/RefinanceOption/LTV"] = {}
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"] = {}
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"]["Cash Out"] = {}
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"]["Cash Out"]["100,000"] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"]["Cash Out"]["100,000"] = new_val
              end
              if r == 169 && cc == 4
                @property_hash["LoanType/PropertyType"] = {}
                @property_hash["LoanType/PropertyType"]["Fixed"] = {}
                @property_hash["LoanType/PropertyType"]["Fixed"]["2nd Home"] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/PropertyType"]["Fixed"]["2nd Home"] = new_val
              end
              if r >= 170 && r <= 173 && cc == 4
                primary_key = value.tr('A-Z% ','')
                @property_hash["LoanType/CLTV"]["Fixed"][primary_key] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/CLTV"]["Fixed"][primary_key] = new_val
              end
              if r == 174 && cc == 4
                @property_hash["LoanType/Term"] = {}
                @property_hash["LoanType/Term"]["Fixed"] = {}
                @property_hash["LoanType/Term"]["Fixed"]["15"] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/Term"]["Fixed"]["15"] = new_val
              end
              # Other Adjustments
              if r == 146 && cc == 14
                @other_adjustment["PropertyType/Term/LTV"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = new_value
              end
              if r == 147 && cc == 14
                @other_adjustment["LoanSize/RefinanceOption"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = new_value
              end
              if r == 148 && cc == 14
                @other_adjustment["LoanSize/RefinanceOption"]["Rate and Term"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["Rate and Term"] = new_value
              end
              # Loans With Secondary Financing
              if value == "Loans With Secondary Financing"
                @secondary_hash["LTV/CLTV/FICO"] = {}
              end

              if r >= 155 && r <= 158 && cc == 12
                ltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key] = {}
              end
              if r >= 155 && r <= 158 && cc == 13
                cltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key] = {}
              end
              if r >= 155 && r <= 158 && cc > 13 && cc <= 15
                ltv_data =  get_value @ltv_data[cc-1]
                ltv_data = ltv_data.tr('() ','')
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = {}
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = value
              end
              # Other Adjustments
              if r == 168 && cc == 12
                @other_adjustment["MiscAdjuster"] = {}
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = new_value
              end
              if r == 169 && cc == 12
                @other_adjustment["LoanAmount"] = {}
                @other_adjustment["LoanAmount"]["0-150,000"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["0-150,000"] = new_value
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
              error_log.save
            end
          end
        end
        adjustment = [@secondary_hash,@cashout_hash,@other_adjustment,@property_hash]
        make_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def olympic_piggyback_high_balance
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Olympic PiggyBack High Balance")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        range1 = 137
        range2 = 172
        @cashout_hash = {}
        @other_adjustment = {}
        @secondary_hash = {}
        @property_hash = {}
        primary_key = ''
        ltv_key = ''
        cltv_key = ''
        ltv_data = ''
        #program
        (51..93).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id
                end

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
                error_log.save
              end

              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  begin
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
                  rescue Exception => e
                    error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet)
                    error_log.save
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

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(153)
          (0..sheet_data.last_column).each do |cc|
            value = sheet_data.cell(r,cc)
            begin
              if value == "LTV / FICO (Terms > 15 years only)"
                first_row = 141
                end_row = 144
                last_column = 8
                first_column = 4
                ltv_row = 140
                num = 1
                ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row,num
              end
              if value == "Cash Out Refinance"
                @cashout_hash["RefinanceOption/FICO/LTV"] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
              end
              if value == "OLYMPIC FIXED 2ND MORTGAGE"
                @property_hash["LoanType/CLTV"] = {}
                @property_hash["LoanType/CLTV"]["Fixed"] = {}
              end
              # Cash Out Refinance
              if r >= 154 && r <= 157 && cc == 4
                primary_key = get_value value
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key] = {}
              end
              if r >= 154 && r <= 157 && cc >= 5 && cc <= 8
                ltv_key = get_value @ltv_data[cc-1]
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = value
              end
              # OLYMPIC FIXED 2ND MORTGAGE
              if r == 166 && cc == 4
                @property_hash["LoanType/RefinanceOption/LTV"] = {}
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"] = {}
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"]["Cash Out"] = {}
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"]["Cash Out"]["100,000"] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"]["Cash Out"]["100,000"] = new_val
              end
              if r == 167 && cc == 4
                @property_hash["LoanType/PropertyType"] = {}
                @property_hash["LoanType/PropertyType"]["Fixed"] = {}
                @property_hash["LoanType/PropertyType"]["Fixed"]["2nd Home"] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/PropertyType"]["Fixed"]["2nd Home"] = new_val
              end
              if r >= 168 && r <= 171 && cc == 4
                primary_key = value.tr('A-Z% ','')
                @property_hash["LoanType/CLTV"]["Fixed"][primary_key] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/CLTV"]["Fixed"][primary_key] = new_val
              end
              if r == 172 && cc == 4
                @property_hash["LoanType/Term"] = {}
                @property_hash["LoanType/Term"]["Fixed"] = {}
                @property_hash["LoanType/Term"]["Fixed"]["15"] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/Term"]["Fixed"]["15"] = new_val
              end
               # Other Adjustments
              if r == 145 && cc == 14
                @other_adjustment["PropertyType/Term/LTV"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = new_value
              end
              if r == 146 && cc == 14
                @other_adjustment["LoanSize/RefinanceOption"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = new_value
              end
              if r == 147 && cc == 14
                @other_adjustment["LoanSize/RefinanceOption"]["Rate and Term"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["Rate and Term"] = new_value
              end
              # Other Adjustments
              if r == 166 && cc == 12
                @other_adjustment["MiscAdjuster"] = {}
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = {}
                cc = cc + 1
                new_val = sheet_data.cell(r,cc)
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = new_val
              end
              if r == 167 && cc == 12
                @other_adjustment["LoanAmount"] = {}
                @other_adjustment["LoanAmount"]["0-100,000"] = {}
                cc = cc + 1
                new_val = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["0-100,000"] = new_val
              end
              # Loans With Secondary Financing
              if value == "Loans With Secondary Financing"
                @secondary_hash["LTV/CLTV/FICO"] = {}
              end
              if r >= 154 && r <= 157 && cc == 12
                ltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key] = {}
              end
              if r >= 154 && r <= 157 && cc == 13
                cltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key] = {}
              end
              if r >= 154 && r <= 157 && cc > 13 && cc <= 15
                ltv_data = get_value @ltv_data[cc-1]
                ltv_data = ltv_data.tr('() ','')
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = {}
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = value
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
              error_log.save
            end
          end
        end
        adjustment = [@secondary_hash,@other_adjustment,@property_hash,@cashout_hash]
        make_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
  end

  def olympic_piggyback_arm
    file = File.join(Rails.root,  'OB_Newfi_Wholesale7019.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Olympic PiggyBack ARM")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        range1 = 139
        range2 = 175
        @cashout_hash = {}
        @other_adjustment = {}
        @secondary_hash = {}
        @property_hash = {}
        primary_key = ''
        ltv_key = ''
        cltv_key = ''
        ltv_data = ''
        #program
        (51..93).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program, sheet
                  @programs_ids << @program.id
                end

                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet)
                error_log.save
              end

              (1..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  begin
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
                  rescue Exception => e
                    error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet)
                    error_log.save
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
        # Adjustments

        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(156)
          (0..sheet_data.last_column).each do |cc|
            begin
              value = sheet_data.cell(r,cc)
              if value == "LTV / FICO (Terms > 15 years only)"
                @other_adjustment["LoanType/LoanSize/LTV"] = {}
                @other_adjustment["LoanType/LoanSize/LTV"]["ARM"] = {}
                @other_adjustment["LoanType/LoanSize/LTV"]["ARM"]["High-Balance"] = {}
                first_row = 141
                end_row = 144
                last_column = 8
                first_column = 4
                ltv_row = 140
                num = 1
                ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row,num
              end
              if value == "Cash Out Refinance"
                @cashout_hash["RefinanceOption/FICO/LTV"] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
              end
              if value == "OLYMPIC FIXED 2ND MORTGAGE"
                @property_hash["LoanType/CLTV"] = {}
                @property_hash["LoanType/CLTV"]["Fixed"] = {}
              end
              # Cash Out Refinance
              if r >= 157 && r <= 160 && cc == 4
                primary_key = get_value value
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key] = {}
              end
              if r >= 157 && r <= 160 && cc >= 5 && cc <= 8
                ltv_key = get_value @ltv_data[cc-1]
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = {}
                @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][ltv_key] = value
              end
              # OLYMPIC FIXED 2ND MORTGAGE
              if r == 169 && cc == 4
                @property_hash["LoanType/RefinanceOption/LTV"] = {}
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"] = {}
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"]["Cash Out"] = {}
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"]["Cash Out"]["100,000"] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/RefinanceOption/LTV"]["Fixed"]["Cash Out"]["100,000"] = new_val
              end
              if r == 170 && cc == 4
                @property_hash["LoanType/PropertyType"] = {}
                @property_hash["LoanType/PropertyType"]["Fixed"] = {}
                @property_hash["LoanType/PropertyType"]["Fixed"]["2nd Home"] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/PropertyType"]["Fixed"]["2nd Home"] = new_val
              end
              if r >= 171 && r <= 174 && cc == 4
                primary_key = value.tr('A-Z% ','')
                @property_hash["LoanType/CLTV"]["Fixed"][primary_key] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/CLTV"]["Fixed"][primary_key] = new_val
              end
              if r == 175 && cc == 4
                @property_hash["LoanType/Term"] = {}
                @property_hash["LoanType/Term"]["Fixed"] = {}
                @property_hash["LoanType/Term"]["Fixed"]["15"] = {}
                cc = cc + 2
                new_val = sheet_data.cell(r,cc)
                @property_hash["LoanType/Term"]["Fixed"]["15"] = new_val
              end
              # Other Adjustments
              if r == 146 && cc == 14
                @other_adjustment["PropertyType/Term/LTV"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = new_value
              end
              if r == 147 && cc == 14
                @other_adjustment["LoanSize/RefinanceOption"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"] = {}
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = new_value
              end
              if r >= 148 && r <= 150 && cc == 14
                if value.include?("<=")
                  ltv_key = get_value value
                else
                  ltv_key = value.tr('A-Za-z%/ ','')
                end
                @other_adjustment["LoanType/LoanSize/LTV"]["ARM"]["High-Balance"][ltv_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanType/LoanSize/LTV"]["ARM"]["High-Balance"][ltv_key] = new_value
              end
              if r == 151 && cc == 14
                @other_adjustment["LoanType/LoanSize/RefinanceOption"] = {}
                @other_adjustment["LoanType/LoanSize/RefinanceOption"]["ARM"] = {}
                @other_adjustment["LoanType/LoanSize/RefinanceOption"]["ARM"]["High-Balance"] = {}
                @other_adjustment["LoanType/LoanSize/RefinanceOption"]["ARM"]["High-Balance"]["Rate and Term"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanType/LoanSize/RefinanceOption"]["ARM"]["High-Balance"]["Rate and Term"] = new_value
              end
              if r == 170 && cc == 12
                @other_adjustment["MiscAdjuster"] = {}
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["MiscAdjuster"]["Escrow Waiver Fee"] = new_value
              end
              if r == 171 && cc == 12
                @other_adjustment["LoanAmount"] = {}
                @other_adjustment["LoanAmount"]["0-150,000"] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment["LoanAmount"]["0-150,000"] = new_value
              end
              # Loans With Secondary Financing
              if value == "Loans With Secondary Financing"
                @secondary_hash["LTV/CLTV/FICO"] = {}
              end
              if r >= 157 && r <= 160 && cc == 12
                ltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key] = {}
              end
              if r >= 157 && r <= 160 && cc == 13
                cltv_key = get_value value
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key] = {}
              end
              if r >= 157 && r <= 160 && cc > 13 && cc <= 15
                ltv_data = get_value @ltv_data[cc-1]
                ltv_data = ltv_data.tr('() ','')
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = {}
                @secondary_hash["LTV/CLTV/FICO"][ltv_key][cltv_key][ltv_data] = value
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
              error_log.save
            end
          end
        end
        adjustment = [@cashout_hash,@secondary_hash,@other_adjustment,@property_hash]
        make_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_newfi_wholesale7019_path(@sheet_obj)
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
          value1 = "0-"+value1.split("<=").last.tr('A-Za-z%$><=/ ','')
        elsif value1.include?(">") || value1.include?("+")
          value1 = value1.split(">").last.tr('^0-9 ', '')+"-Inf"
        else
          value1 = value1.tr('% ','')
        end
      end
    end

    def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end

    def get_program
      @program = Program.find(params[:id])
    end

    def program_property value1, sheet
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

       # Arm Basic
      if @program.program_name.include?("3/1") || @program.program_name.include?("3 / 1")
        arm_basic = 3
      elsif @program.program_name.include?("5/1") || @program.program_name.include?("5 / 1")
        arm_basic = 5
      elsif @program.program_name.include?("7/1") || @program.program_name.include?("7 / 1")
        arm_basic = 7
      elsif @program.program_name.include?("10/1") || @program.program_name.include?("10 / 1")
        arm_basic = 10          
      end

      # High Balance
      jumbo_high_balance = false
      if @program.program_name.include?("High Bal")
        jumbo_high_balance = true
      end

      # Arm Advanced
      if @program.program_name.include?("ARM")
        arm_advanced = @program.program_name.split.last
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
      @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline, sheet_name: sheet, jumbo_high_balance: jumbo_high_balance, arm_basic: arm_basic, arm_advanced: arm_advanced)
    end
    
    def make_adjust(block_hash, sheet)
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

    def ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row,num
      @adjustment_hash = {}
      primary_key = ''
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
                if value == "LTV / FICO (Terms > 15 years only)"
                  primary_key = "LoanType/Term/LTV/FICO"
                  @adjustment_hash["Term/FICO/LTV"] = {}
                  @adjustment_hash["Term/FICO/LTV"]["15-Inf"] = {}
                end
                if r >= first_row && r <= end_row && cc == first_column
                  ltv_key = get_value value
                  @adjustment_hash["Term/FICO/LTV"]["15-Inf"][ltv_key] = {}
                end
                if r >= first_row && r <= end_row && cc > first_column && cc <= last_column
                  cltv_key = get_value @ltv_data[cc-num]
                  @adjustment_hash["Term/FICO/LTV"]["15-Inf"][ltv_key][cltv_key] = {}
                  @adjustment_hash["Term/FICO/LTV"]["15-Inf"][ltv_key][cltv_key] = value
                end
              end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, sheet_name: sheet)
              error_log.save
            end
          end
        end
      end
      adjustment = [@adjustment_hash]
      make_adjust(adjustment,sheet)
      # create_program_association_with_adjustment(sheet)
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
