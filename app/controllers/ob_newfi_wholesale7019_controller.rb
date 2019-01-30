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
reprograms_ob_newfi_wholesale7019_path
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
          if ((row.compact.count > 1) && (row.compact.count < 4)) || row.include?("7/1 LIBOR ARM BISCAYNE JUMBO") || row.include?("10/1 LIBOR ARM BISCAYNE JUMBO")
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
                    elsif (c_i == 1)
                      @block_hash[main_key][key][30] = value
                    elsif (c_i == 2)
                      @block_hash[main_key][key][45] = value
                    elsif (c_i == 3)
                      @block_hash[main_key][key][60] = value
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
        create_program_association_with_adjustment(sheet)
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
        (77..113).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(81)
          if row.compact.count > 0
            (0..12).each do |cc|
              value = sheet_data.cell(r,cc)
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

        # Adjustments
        (101..120).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(104)
          if row.compact.count >= 1
            (0..19).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "FICO - Loan Amount "
                  primary_key = "LoanAmount/LTV/FICO"
                  @adjustment_hash[primary_key] = {}
                end
                # FICO - Loan Amount 
                if r >= 105 && r <= 110 && cc == 5
                  secondary_key = value
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 105 && r <= 110 && cc > 5 && cc <= 8
                  ltv_key = get_value @ltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = value
                end
                # Other Adjustments
                if r == 115 && cc == 14
                  primary_key = "Escrow Waiver Fee"
                  @other_adjustment[primary_key] = {}
                  if @other_adjustment[primary_key] = {}
                    cc = cc + 4
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment[primary_key] = new_value
                  end
                end
                if r == 116 && cc == 14
                  primary_key = "LoanAmount"
                  secondary_key = "0<$150,000"
                  @other_adjustment[primary_key] = {}
                  @other_adjustment[primary_key][secondary_key] = {}
                  if @other_adjustment[primary_key][secondary_key] = {}
                    cc = cc + 4
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment[primary_key][secondary_key] = new_value
                  end
                end
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
        @price_adjustment = {}
        @family_adjustment = {}
        @condo_adjustment = {}
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

        # Adjustments
        (101..128).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(104)
          @cltv_data = sheet_data.row(114)
          (0..16).each do |cc|
            value = sheet_data.cell(r,cc)
            if value.present?
              if value == " Price Adjustments"
                primary_key = "LoanType/LTV/FICO"
                @price_adjustment[primary_key] = {}
              end
              if value == "Multi Family 2- 4 Unit LTV/FICO Adjusters"
                primary_key = "2-4 Unit"
                secondary_key = "FICO/LTV"
                @family_adjustment[primary_key] = {}
                @family_adjustment[primary_key][secondary_key] = {}
              end
              if value == "Condo LTV/FICO Adjusters"
                primary_key1 = "Condo"
                secondary_key1 = "LTV/FICO"
                @condo_adjustment[primary_key1] = {}
                @condo_adjustment[primary_key1][secondary_key1] = {}
              end
              if value == "HIGH BALANCE"
                primary_key = "HighBalance"
                secondary_key = "LTV/FICO"
                @high_adjustment[primary_key] = {}
              end
              if r >= 105 && r <= 112 && cc == 7
                secondary_key = value
                @price_adjustment[primary_key][secondary_key] = {}
              end
              if r >= 105 && r <= 112 && cc > 7 && cc <= 112
                ltv_key = @ltv_data[cc-1]
                @price_adjustment[primary_key][secondary_key][ltv_key] = {}
                @price_adjustment[primary_key][secondary_key][ltv_key] = value
              end
              # Multi Family 2- 4 Unit LTV/FICO Adjusters
              if r >= 115 && r <= 122 && cc == 6
                ltv_key = value
                @family_adjustment[primary_key][secondary_key][ltv_key] = {}
              end
              if r >= 115 && r <= 122 && cc > 6 && cc <= 10
                cltv_key = get_value @cltv_data[cc-1]
                @family_adjustment[primary_key][secondary_key][ltv_key][cltv_key] = {}
                @family_adjustment[primary_key][secondary_key][ltv_key][cltv_key] = value
              end
              # Condo LTV/FICO Adjusters
              if r >= 115 && r <= 122 && cc == 12
                ltv_key1 = value
                @condo_adjustment[primary_key1][secondary_key1][ltv_key1] = {}
              end
              if r >= 115 && r <= 122 && cc > 12 && cc <= 16
                cltv_key1 = get_value @cltv_data[cc-1]
                @condo_adjustment[primary_key1][secondary_key1][ltv_key1][cltv_key1] = {}
                @condo_adjustment[primary_key1][secondary_key1][ltv_key1][cltv_key1] = value
              end
              # HIGH BALANCE
              if r >= 124 && r <= 126 && cc == 9
              end
            end
          end
        end
        adjustment = [@price_adjustment,@family_adjustment,condo_adjustment]
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

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(117)
          (0..sheet_data.last_column).each do |cc|
            value = sheet_data.cell(r,cc)
            if value == "LTV / FICO (Terms > 15 years only)"
              first_row = 105
              end_row = 108
              last_column = 13
              first_column = 5
              ltv_row = 104
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
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
              ltv_data = get_value @ltv_data[cc-1]
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = {}
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = value
            end
            # Other Adjustments
            if r == 108 && cc == 15
              primary_key = value
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 109 && cc == 15
              primary_key = "PropertyType/LTV/Term"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 110 && cc == 15
              primary_key = "RefinanceOption/HighBalance"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 111 && cc == 15
              primary_key = "PropertyType/HighBalance"
              secondary_key = "0 < 75%"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
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

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(153)
          (0..sheet_data.last_column).each do |cc|
            value = sheet_data.cell(r,cc)
            if value == "LTV / FICO (Terms > 15 years only)"
              first_row = 141
              end_row = 147
              last_column = 12
              first_column = 4
              ltv_row = 140
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end
            if value == "Cash Out Refinance"
              first_row = 154
              end_row = 160
              first_column = 4
              last_column = 8
              ltv_row = 153
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end

            # Other Adjustments
            if r == 141 && cc == 14
              primary_key = "PropertyType/LTV"
              secondary_key = "0<= 75"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end

            if r == 142 && cc == 14
              secondary_key = "> 75 and <=80"
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 143 && cc == 14
              secondary_key = ">80"
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 144 && cc == 14
              primary_key = value
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 145 && cc == 14
              primary_key = "PropertyType/LTV/Term"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 146 && cc == 14
              primary_key = "RefinanceOption/HighBalance"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 147 && cc == 14
              primary_key = "PropertyType/HighBalance/LTV/CLTV"
              secondary_key = "0 < 75%"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 166 && cc == 13
              primary_key = "Escrow Waiver Fee"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 167 && cc == 13
              primary_key = "Loan Amount"
              secondary_key = "0< $150,000"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 168 && cc == 13
              secondary_key = "Non CA Conforming >= $250k "
              if @other_adjustment[primary_key].present?
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 169 && cc == 13
              secondary_key = "Non CA Conforming >= $200k < $250k"
              if @other_adjustment[primary_key].present?
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 170 && cc == 13
              primary_key = "FICO"
              secondary_key =  "0 < 680"
              if @other_adjustment[primary_key] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 170 && cc == 13
              primary_key = "Loan Amount"
              secondary_key =  ">= $275k"
              if @other_adjustment[primary_key] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 170 && cc == 13
              secondary_key =  ">= $200k < $275k"
              if @other_adjustment[primary_key].present?
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            # Loans With Secondary Financing
            if value == "Loans With Secondary Financing"
              primary_key = "LTV/CLTV/FICO"
              @secondary_hash[primary_key] = {}
            end
            if r >= 154 && r <= 158 && cc == 12
              ltv_key = get_value value
              @secondary_hash[primary_key][ltv_key] = {}
            end
            if r >= 154 && r <= 158 && cc == 13
              cltv_key = get_value value
              @secondary_hash[primary_key][ltv_key][cltv_key] = {}
            end
            if r >= 154 && r <= 158 && cc > 13 && cc <= 15
              ltv_data = get_value @ltv_data[cc-1]
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = {}
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = value
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

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(153)
          (0..sheet_data.last_column).each do |cc|
            value = sheet_data.cell(r,cc)
            if value == "LTV / FICO (Terms > 15 years only)"
              first_row = 141
              end_row = 147
              last_column = 12
              first_column = 4
              ltv_row = 140
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end
            if value == "Cash Out Refinance"
              first_row = 154
              end_row = 160
              first_column = 4
              last_column = 8
              ltv_row = 153
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end

            # Other Adjustments
            if r == 141 && cc == 14
              primary_key = "PropertyType/LTV"
              secondary_key = "0<= 75"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end

            if r == 142 && cc == 14
              secondary_key = "> 75 and <=80"
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 143 && cc == 14
              secondary_key = ">80"
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 144 && cc == 14
              primary_key = value
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 145 && cc == 14
              primary_key = "PropertyType/LTV/Term"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 146 && cc == 14
              primary_key = "RefinanceOption/HighBalance"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 147 && cc == 14
              primary_key = "PropertyType/HighBalance/LTV/CLTV"
              secondary_key = "0 < 75%"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 166 && cc == 14
              primary_key = "Escrow Waiver Fee"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 167 && cc == 14
              primary_key = "Loan Amount"
              secondary_key = "0< $150k"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 168 && cc == 14
              secondary_key = "Standard Conforming >= $300k "
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 169 && cc == 14
              secondary_key = "Non CA Conforming >= $200k < $300k"
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
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
              primary_key = "LTV/CLTV/FICO"
              @secondary_hash[primary_key] = {}
            end
            if r >= 154 && r <= 158 && cc == 12
              ltv_key = get_value value
              @secondary_hash[primary_key][ltv_key] = {}
            end
            if r >= 154 && r <= 158 && cc == 13
              cltv_key = get_value value
              @secondary_hash[primary_key][ltv_key][cltv_key] = {}
            end
            if r >= 154 && r <= 158 && cc > 13 && cc <= 15
              ltv_data = get_value @ltv_data[cc-1]
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = {}
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = value
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

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(156)
          (0..sheet_data.last_column).each do |cc|
            value = sheet_data.cell(r,cc)
            if value == "LTV / FICO (Terms > 15 years only)"
              first_row = 141
              end_row = 147
              last_column = 12
              first_column = 4
              ltv_row = 140
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end
            if value == "Cash out Refinance"
              first_row = 155
              end_row = 161
              first_column = 4
              last_column = 8
              ltv_row = 154
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end
            # Other Adjustments
            if r == 141 && cc == 14
              primary_key = "PropertyType/LTV"
              secondary_key = "0<= 75"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end

            if r == 142 && cc == 14
              secondary_key = "> 75 and <=80"
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 143 && cc == 14
              secondary_key = ">80"
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 144 && cc == 14
              primary_key = value
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 145 && cc == 14
              primary_key = "PropertyType/LTV/Term"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 146 && cc == 14
              primary_key = "RefinanceOption/HighBalance"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 147 && cc == 14
              primary_key = "PropertyType/HighBalance/LTV/CLTV"
              secondary_key = "0 < 75%"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 148 && cc == 14
              secondary_key = "75.01% - 90%"
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 149 && cc == 14
              secondary_key = "90.01% - 95%"
              # debugger
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 150 && cc == 14
              primary_key = "LoanType/HighBalance"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 165 && cc == 12
              primary_key = "Escrow Waiver Fee"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 166 && cc == 12
              primary_key = "Loan Amount"
              secondary_key = "0< $150,000"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 167 && cc == 12
              secondary_key = "CA Conforming >= $250k "
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 168 && cc == 12
              secondary_key = "Non CA Conforming >= $200k < $250k"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 169 && cc == 12
              primary_key = "FICO"
              secondary_key =  "0<680"
              if @other_adjustment[primary_key] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 170 && cc == 12
              primary_key = "Loan Amount"
              secondary_key =  ">= $275k"
              if @other_adjustment[primary_key] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 170 && cc == 12
              secondary_key =  ">= $275k"
              if @other_adjustment[primary_key] = {}
                cc = cc + 4
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end

            # Loans With Secondary Financing
            if value == "Loan With Secondary Financing"
              primary_key = "LTV/CLTV/FICO"
              @secondary_hash[primary_key] = {}
            end
            if r >= 155 && r <= 159 && cc == 12
              ltv_key = value
              @secondary_hash[primary_key][ltv_key] = {}
            end
            if r >= 155 && r <= 159 && cc == 13
              cltv_key = value
              @secondary_hash[primary_key][ltv_key][cltv_key] = {}
            end
            if r >= 155 && r <= 159 && cc > 13 && cc <= 15
              ltv_data =  @ltv_data[cc-1]
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = {}
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = value
            end
          end
        end
        adjustment = [@secondary_hash,@cashout_hash,@other_adjustment]
        make_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
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

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(154)
          (0..sheet_data.last_column).each do |cc|
            value = sheet_data.cell(r,cc)
            if value == "LTV / FICO (Terms > 15 years only)"
              first_row = 141
              end_row = 144
              last_column = 8
              first_column = 4
              ltv_row = 140
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end
            if value == "Cash Out Refinance"
              first_row = 155
              end_row = 158
              first_column = 4
              last_column = 8
              ltv_row = 154
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end
            if value == "OLYMPIC FIXED 2ND MORTGAGE"
              primary_key = "LoanType/LTV"
              @cashout_hash[primary_key] = {}
            end
            # OLYMPIC FIXED 2ND MORTGAGE
            if r >= 166 && r <= 174 && cc == 4 && r != 167
              secondary_key = value
              @cashout_hash[primary_key][secondary_key] = {}
              if @cashout_hash[primary_key][secondary_key] == {}
                cc = cc +2
                new_value = sheet_data.cell(r,cc)
                @cashout_hash[primary_key][secondary_key] = new_value
              end              
            end
            # Other Adjustments
            if r == 146 && cc == 14
              primary_key = "PropertyType/LTV/Term"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 147 && cc == 14
              primary_key = "RefinanceOption/HighBalance"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 148 && cc == 14
              primary_key = "LoanType/HighBalance"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 168 && cc == 12
              primary_key1 = "Escrow Waiver Fee"
              @other_adjustment[primary_key1] = {}
              if @other_adjustment[primary_key1] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key1] = new_value
              end
            end
            if r == 169 && cc == 12
              primary_key1 = "Loan Amount 0< $150k"
              @other_adjustment[primary_key1] = {}
              if @other_adjustment[primary_key1] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key1] = new_value
              end
            end
            # Loans With Secondary Financing
            if value == "Loans With Secondary Financing"
              primary_key = "LTV/CLTV/FICO"
              @secondary_hash[primary_key] = {}
            end

            if r >= 155 && r <= 158 && cc == 12
              ltv_key = value
              @secondary_hash[primary_key][ltv_key] = {}
            end
            if r >= 155 && r <= 158 && cc == 13
              cltv_key = value
              @secondary_hash[primary_key][ltv_key][cltv_key] = {}
            end
            if r >= 155 && r <= 158 && cc > 13 && cc <= 15
              ltv_data =  @ltv_data[cc-1]
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = {}
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = value
            end
          end
        end
        adjustment = [@secondary_hash,@other_adjustment,@other_adjustment]
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

        # Adjustments
        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(153)
          (0..sheet_data.last_column).each do |cc|
            value = sheet_data.cell(r,cc)
            if value == "LTV / FICO (Terms > 15 years only)"
              first_row = 141
              end_row = 144
              last_column = 8
              first_column = 4
              ltv_row = 140
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end
            if value == "Cash Out Refinance"
              first_row = 154
              end_row = 157
              first_column = 4
              last_column = 8
              ltv_row = 153
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end
            if value == "OLYMPIC FIXED 2ND MORTGAGE"
              primary_key = "LoanType/LTV"
              @cashout_hash[primary_key] = {}
            end
            # OLYMPIC FIXED 2ND MORTGAGE
            if r >= 164 && r <= 172 && cc == 4 && r != 165
              secondary_key = value
              @cashout_hash[primary_key][secondary_key] = {}
              if @cashout_hash[primary_key][secondary_key] == {}
                cc = cc +2
                new_value = sheet_data.cell(r,cc)
                @cashout_hash[primary_key][secondary_key] = new_value
              end              
            end
            # Other Adjustments
            if r == 145 && cc == 14
              primary_key = "PropertyType/LTV/Term"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 146 && cc == 14
              primary_key = "RefinanceOption/HighBalance"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 147 && cc == 14
              primary_key = "LoanType/HighBalance"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 170 && cc == 12
              primary_key1 = "Escrow Waiver Fee"
              @other_adjustment[primary_key1] = {}
              if @other_adjustment[primary_key1] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key1] = new_value
              end
            end
            if r == 171 && cc == 12
              primary_key1 = "Loan Amount 0< $100k"
              @other_adjustment[primary_key1] = {}
              if @other_adjustment[primary_key1] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key1] = new_value
              end
            end
            # Loans With Secondary Financing
            if value == "Loans With Secondary Financing"
              primary_key = "LTV/CLTV/FICO"
              @secondary_hash[primary_key] = {}
            end

            if r >= 154 && r <= 157 && cc == 12
              ltv_key = value
              @secondary_hash[primary_key][ltv_key] = {}
            end
            if r >= 154 && r <= 157 && cc == 13
              cltv_key = value
              @secondary_hash[primary_key][ltv_key][cltv_key] = {}
            end
            if r >= 154 && r <= 157 && cc > 13 && cc <= 15
              ltv_data =  @ltv_data[cc-1]
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = {}
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = value
            end
          end
        end
        adjustment = [@secondary_hash,@other_adjustment,@other_adjustment]
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
        # Adjustments

        (range1..range2).each do |r|
          @ltv_data = sheet_data.row(156)
          (0..sheet_data.last_column).each do |cc|
            value = sheet_data.cell(r,cc)
            if value == "LTV / FICO (Terms > 15 years only)"
              first_row = 141
              end_row = 144
              last_column = 8
              first_column = 4
              ltv_row = 140
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end
            if value == "Cash Out Refinance"
              first_row = 157
              end_row = 160
              first_column = 4
              last_column = 8
              ltv_row = 156
              ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
            end
            if value == "OLYMPIC FIXED 2ND MORTGAGE"
              primary_key = "LoanType/LTV"
              @cashout_hash[primary_key] = {}
            end
            # OLYMPIC FIXED 2ND MORTGAGE
            if r >= 167 && r <= 175 && cc == 4 && r != 168
              secondary_key = value
              @cashout_hash[primary_key][secondary_key] = {}
              if @cashout_hash[primary_key][secondary_key] == {}
                cc = cc +2
                new_value = sheet_data.cell(r,cc)
                @cashout_hash[primary_key][secondary_key] = new_value
              end              
            end
            # Other Adjustments
            if r == 146 && cc == 14
              primary_key = "PropertyType/LTV/Term"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 147 && cc == 14
              primary_key = "RefinanceOption/HighBalance"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 148 && cc == 14
              primary_key = "PropertyType/HighBalance/LTV/CLTV"
              secondary_key = "0 < 75%"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 149 && cc == 14
              secondary_key = "75.01% - 90%"
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 150 && cc == 14
              secondary_key = "90.01% - 95%"
              if @other_adjustment[primary_key].present?
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key][secondary_key] = {}
                @other_adjustment[primary_key][secondary_key] = new_value
              end
            end
            if r == 151 && cc == 14
              primary_key = "LoanType/HighBalance"
              @other_adjustment[primary_key] = {}
              if @other_adjustment[primary_key] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key] = new_value
              end
            end
            if r == 170 && cc == 12
              primary_key1 = "Escrow Waiver Fee"
              @other_adjustment[primary_key1] = {}
              if @other_adjustment[primary_key1] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key1] = new_value
              end
            end
            if r == 171 && cc == 12
              primary_key1 = "Loan Amount 0< $150k"
              @other_adjustment[primary_key1] = {}
              if @other_adjustment[primary_key1] = {}
                cc = cc + 1
                new_value = sheet_data.cell(r,cc)
                @other_adjustment[primary_key1] = new_value
              end
            end
            # Loans With Secondary Financing
            if value == "Loans With Secondary Financing"
              primary_key = "LTV/CLTV/FICO"
              @secondary_hash[primary_key] = {}
            end

            if r >= 157 && r <= 160 && cc == 12
              ltv_key = get_value value
              @secondary_hash[primary_key][ltv_key] = {}
            end
            if r >= 157 && r <= 160 && cc == 13
              cltv_key = get_value value
              @secondary_hash[primary_key][ltv_key][cltv_key] = {}
            end
            if r >= 157 && r <= 160 && cc > 13 && cc <= 15
              ltv_data = get_value @ltv_data[cc-1]
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = {}
              @secondary_hash[primary_key][ltv_key][cltv_key][ltv_data] = value
            end
          end
        end
        adjustment = [@secondary_hash,@secondary_hash,@other_adjustment]
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
        if value1.include?("FICO <")
          value1 = "0"+value1.split("FICO").last
        elsif value1.include?("<=") || value1.include?(">=") || value1.include?("<")
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
        Adjustment.create(data: hash,sheet_name: sheet)
      end
    end

    def ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
      @adjustment_hash = {}
      primary_key = ''
      ltv_key = ''
      cltv_key = ''
      (range1..range2).each do |r|
        row = sheet_data.row(r)
        @ltv_data = sheet_data.row(ltv_row)
        if row.compact.count >= 1
          (0..last_column).each do |cc|
            value = sheet_data.cell(r,cc)
            if value.present?
              if value == "LTV / FICO (Terms > 15 years only)"
                primary_key = "LoanType/Term/LTV/FICO" 
                @adjustment_hash[primary_key] = {}
              end
              if value == "Cash Out Refinance"
                primary_key = "RefinanceOption/FICO/LTV" 
                @adjustment_hash[primary_key] = {}
              end
              if r >= first_row && r <= end_row && cc == first_column
                ltv_key = value
                @adjustment_hash[primary_key][ltv_key] = {}
              end
              if r >= first_row && r <= end_row && cc > first_column && cc <= 8
                cltv_key = get_value @ltv_data[cc-1]
                @adjustment_hash[primary_key][ltv_key][cltv_key] = {}
                @adjustment_hash[primary_key][ltv_key][cltv_key] = value
              end
            end
          end
        end
      end
      # @adjustment_hash.keys.each do |a_hash|
      #   if @adjustment_hash[a_hash] == {}
      #     @adjustment_hash.shift
      #   end
      # end
      # return @adjustment_hash
      adjustment = [@adjustment_hash]
      make_adjust(adjustment,sheet)
      create_program_association_with_adjustment(sheet)
    end

    # def subordinate_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row, cltv_column
    #   @subordinate_hash = {}
    #   primary_key = ''
    #   ltv_key = ''
    #   cltv_key = ''
    #   (range1..range2).each do |r|
    #     row = sheet_data.row(r)
    #     @ltv_data = sheet_data.row(ltv_row)
    #     if row.compact.count >= 1
    #       (0..last_column).each do |cc|
    #         value = sheet_data.cell(r,cc)
    #         if value.present?
    #           primary_key = "LTV/CLTV/FICO" 
    #           @subordinate_hash[primary_key] = {}
    #           if r >= first_row && r <= end_row && cc == first_column
    #             ltv_key = get_value value
    #             @subordinate_hash[primary_key][ltv_key] = {}
    #             if @subordinate_hash[primary_key][ltv_key] = {}
    #               new_value = sheet_data.cell(r,cltv_column)
    #               @subordinate_hash[primary_key][ltv_key][new_value] = {}
    #             end
    #           end
    #           # if r >= first_row && r <= end_row && cc == cltv_column
    #           #   cltv_key = get_value value
    #           #   debugger
    #           #   @subordinate_hash[primary_key][ltv_key][cltv_key] = {}
    #           # end
    #           # if r >= first_row && r <= end_row && cc > cltv_column && cc <= last_column
    #           #   debugger
    #           #   cltv_data = get_value @ltv_data[cc-1]
    #           #   @subordinate_hash[primary_key][ltv_key][cltv_key] = {}
    #           #   @subordinate_hash[primary_key][ltv_key][cltv_key] = value
    #           # end
    #         end
    #       end
    #     end
    #   end 
    # end

    def create_program_association_with_adjustment(sheet)
      adjustment_list = Adjustment.where(sheet_name: sheet)
      adjustment_list.each_with_index do |adj_ment, index|
        key_list = adj_ment.data.keys.first.split("/")
        program_filter1={}
        program_filter2={}

        if key_list.present?
          key_list.each_with_index do |key_name, key_index|
            if key_name == "LoanType" || key_name == "Term"
              program_filter1[key_name.underscore] = nil
            end

            if key_name == "FICO"
            end

            if key_name == "LTV"
            end

            if key_name == "LoanAmount"
            end

            if key_name == "FinancingType"
            end

            if key_name == "CashOut"
            end
          end

          program_list1 = Program.where.not(program_filter1)
          program_list2 = program_list1.where(program_filter2)

          if program_list2.present?
            program_list2.each do |program|
              program.adjustments.destroy_all
            end

            program_list2.each do |program|
              program.adjustments << adj_ment
            end
          end
        end
      end
    end
  end