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
        (53..92).each do |r|
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
        # adjustments
        (108..153).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(109)
          @cltv_data = sheet_data.row(136)
          (0..13).each do |cc|
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
          end
        end
        adjustment = [@adjustment_hash,@purpose_adjustment,@highAdjustment]
        make_adjust(adjustment,sheet)
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
        f_key = ''
        first_key = ''
        cc = ''
        ccc = ''
        c_val = ''

        #program
        (51..92).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              @title = sheet_data.cell(r,cc)
              if @title.present?
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
        f_key = ''
        first_key = ''
        cc = ''
        ccc = ''
        c_val = ''

        #program
        (51..92).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              @title = sheet_data.cell(r,cc)
              if @title.present?
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
        f_key = ''
        first_key = ''
        cc = ''
        ccc = ''
        c_val = ''

        #program
        (51..92).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              @title = sheet_data.cell(r,cc)
              if @title.present?
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
        f_key = ''
        first_key = ''
        cc = ''
        ccc = ''
        c_val = ''

        #program
        (51..93).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              @title = sheet_data.cell(r,cc)
              if @title.present?
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
        f_key = ''
        first_key = ''
        cc = ''
        ccc = ''
        c_val = ''

        #program
        (51..115).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              @title = sheet_data.cell(r,cc)
              if @title.present?
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
        f_key = ''
        first_key = ''
        cc = ''
        ccc = ''
        c_val = ''

        #program
        (51..92).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 5*max_column + (3+max_column) # 3 / 9 / 15
              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != "5/1 CMT ARM 1/1/5 VA"
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
        f_key = ''
        first_key = ''
        cc = ''
        ccc = ''
        c_val = ''

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
        cc = ''
        ccc = ''
        c_val = ''

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
        cc = ''
        ccc = ''
        c_val = ''

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
        cc = ''
        ccc = ''
        c_val = ''

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
        cc = ''
        ccc = ''
        c_val = ''

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
        cc = ''
        ccc = ''
        c_val = ''

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
        cc = ''
        ccc = ''
        c_val = ''

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
    def make_adjust(block_hash, sheet)
      block_hash.each do |hash|
        Adjustment.create(data: hash.to_json,sheet_name: sheet)
      end
    end
end