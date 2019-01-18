class ObCardinalFinancialWholesale10742Controller < ApplicationController
	before_action :get_sheet, only: [:programs, :ak]
	before_action :get_program, only: [:single_program, :program_property]
	def index
		file = File.join(Rails.root,  'OB_Cardinal_Financial_Wholesale10742.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "AK")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "Cardinal Financial"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
	end

  def ak
    file = File.join(Rails.root,  'OB_Cardinal_Financial_Wholesale10742.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "AK")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (71..298).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4)) && !row.include?("Specials") && !row.include?("Max Net Rebate") && !row.include?("State Adjustment") && !row.include?("ALASKA") && !row.include?("> 484.35K") && !row.include?("7. Not applicable if the subordinate financing is an Affordable Second") && !row.include?("> 100k & ≤ 125k ") && !row.include?("8. Applies to Relief Refiance Mortgages Only") && !row.include?("*Approval Required by Credit Committee for all No Credit loans") && !row.include?("Other Specific Adjustments") && !row.include?("*State Adj. Applied After The Cap") && !row.include?("104.5") || (row.include?("Jumbo 5/1 ARM"))
            rr = r + 1 
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|
              index = index +1
              cc = 1 + max_column*10 + index# (2 / 13 / 24 / 35)
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
                (0..8).each_with_index do |index, c_i|
                  rrr = rr + max_row +1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*(c_i/2)
                        @program.save
                      end
                      @block_hash[main_key][key][15*(c_i/2)] = value unless @block_hash[main_key][key].nil?
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
        # Freddie programs
        (458..684).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4)) && !row.include?("Specials") && !row.include?("Max Net Rebate") && !row.include?("State Adjustment") && !row.include?("ALASKA") && !row.include?("> 484.35K") && !row.include?("7. Not applicable if the subordinate financing is an Affordable Second") && !row.include?("> 100k & ≤ 125k ") && !row.include?("8. Applies to Relief Refiance Mortgages Only") && !row.include?("*Approval Required by Credit Committee for all No Credit loans") && !row.include?("Other Specific Adjustments") && !row.include?("*State Adj. Applied After The Cap") && !row.include?("104.5") || (row.include?("Jumbo 5/1 ARM"))
            rr = r + 1 
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|
              index = index +1
              cc = 1 + max_column*10 + index# (2 / 13 / 24 / 35)
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
                (0..8).each_with_index do |index, c_i|
                  rrr = rr + max_row +1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*(c_i/2)
                        @program.save
                      end
                      @block_hash[main_key][key][15*(c_i/2)] = value unless @block_hash[main_key][key].nil?
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
        # FHA Va Usda programs
        (844..1006).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4)) && !row.include?("Specials") && !row.include?("Max Net Rebate") && !row.include?("State Adjustment") && !row.include?("ALASKA") && !row.include?("> 484.35K") && !row.include?("7. Not applicable if the subordinate financing is an Affordable Second") && !row.include?("> 100k & ≤ 125k ") && !row.include?("8. Applies to Relief Refiance Mortgages Only") && !row.include?("*Approval Required by Credit Committee for all No Credit loans") && !row.include?("Other Specific Adjustments") && !row.include?("*State Adj. Applied After The Cap") && !row.include?("104.5") || (row.include?("Jumbo 5/1 ARM"))
            rr = r + 1 
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|
              index = index +1
              cc = 1 + max_column*10 + index# (2 / 13 / 24 / 35)
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
                (0..8).each_with_index do |index, c_i|
                  rrr = rr + max_row +1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*(c_i/2)
                        @program.save
                      end
                      @block_hash[main_key][key][15*(c_i/2)] = value unless @block_hash[main_key][key].nil?
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
        # Non Conforming programs
        (1126..1145).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4)) && !row.include?("Specials") && !row.include?("Max Net Rebate") && !row.include?("State Adjustment") && !row.include?("ALASKA") && !row.include?("> 484.35K") && !row.include?("7. Not applicable if the subordinate financing is an Affordable Second") && !row.include?("> 100k & ≤ 125k ") && !row.include?("8. Applies to Relief Refiance Mortgages Only") && !row.include?("*Approval Required by Credit Committee for all No Credit loans") && !row.include?("Other Specific Adjustments") && !row.include?("*State Adj. Applied After The Cap") && !row.include?("104.5") || (row.include?("Jumbo 5/1 ARM"))
            rr = r + 1 
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|
              index = index +1
              cc = 1 + max_column*10 + index# (2 / 13 / 24 / 35)
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
                (0..8).each_with_index do |index, c_i|
                  rrr = rr + max_row +1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*(c_i/2)
                        @program.save
                      end
                      @block_hash[main_key][key][15*(c_i/2)] = value unless @block_hash[main_key][key].nil?
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
        # Jumbo Non Conforming programs
        (1220..1260).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4)) && !row.include?("Specials") && !row.include?("Max Net Rebate") && !row.include?("State Adjustment") && !row.include?("ALASKA") && !row.include?("> 484.35K") && !row.include?("7. Not applicable if the subordinate financing is an Affordable Second") && !row.include?("> 100k & ≤ 125k ") && !row.include?("8. Applies to Relief Refiance Mortgages Only") && !row.include?("*Approval Required by Credit Committee for all No Credit loans") && !row.include?("Other Specific Adjustments") && !row.include?("*State Adj. Applied After The Cap") && !row.include?("104.5") || (row.include?("Jumbo 5/1 ARM"))
            rr = r + 1 
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|
              index = index +1
              cc = 1 + max_column*10 + index# (2 / 13 / 24 / 35)
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
                (0..8).each_with_index do |index, c_i|
                  rrr = rr + max_row +1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*(c_i/2)
                        @program.save
                      end
                      @block_hash[main_key][key][15*(c_i/2)] = value unless @block_hash[main_key][key].nil?
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
    end
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
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
end
