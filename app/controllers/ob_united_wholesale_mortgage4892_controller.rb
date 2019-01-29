class ObUnitedWholesaleMortgage4892Controller < ApplicationController
	before_action :get_sheet, only: [:programs, :conv, :govt, :govt_arms, :non_conf]
  before_action :read_sheet, only: [:conv, :govt, :govt_arms, :non_conf]
  before_action :get_program, only: [:single_program, :program_property]

	def index
    file = File.join(Rails.root,'OB_United_Wholesale_Mortgage4892.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "Conv")
          # headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "UWM United Wholesale Mortgage"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

	def conv
    @xlsx.sheets.each do |sheet|
      if (sheet == "Conv")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        start_range = 13
        end_range = 103
        row_count = 15
        column_count = 3
        num1 = 2
        num2 = 4
        make_program start_range, end_range, sheet_data, row_count, column_count, num1, num2
      end
    end
    redirect_to programs_ob_united_wholesale_mortgage4892_path(@sheet_obj)
	end

  def govt
    @xlsx.sheets.each do |sheet|
      if (sheet == "Govt")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        start_range = 9
        end_range = 93
        row_count = 15
        column_count = 3
        num1 = 1
        num2 = 4
        make_program start_range, end_range, sheet_data, row_count, column_count, num1, num2
      end
    end
    redirect_to programs_ob_united_wholesale_mortgage4892_path(@sheet_obj)
  end

  def govt_arms
    @xlsx.sheets.each do |sheet|
      if (sheet == "Govt ARMS")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @loan_amount = {}
        @other_adjustment = {}
        primary_key = ''
        secondary_key = ''
        ltv_key = ''
        #program
        (9..14).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1 # 3 / 7 / 11
              @title = sheet_data.cell(r,cc)
              if @title.present? 
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property
              end
              @block_hash = {}
              key = ''
              (1..4).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
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
        (15..28).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4)) 
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1 # 3 / 7 / 11
              @title = sheet_data.cell(r,cc)
              if @title.present? 
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property
              end
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..11).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
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
        (30..35).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1 # 3 / 7 / 11
              @title = sheet_data.cell(r,cc)
              if @title.present? 
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property
              end
              @block_hash = {}
              key = ''
              (1..4).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
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
        (36..49).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 4)) 
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1 # 3 / 7 / 11
              @title = sheet_data.cell(r,cc)
              if @title.present? 
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property
              end
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..11).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
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

        # adjustments
        (51..66).each do |r|
          row = sheet_data.row(r)
          if row.compact.count >= 1
            (0..13).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "GOVERNMENT PRICE ADJUSTMENTS"
                  primary_key = "LoanAmount"
                  @adjustment_hash[primary_key] = {}
                end
                if value == "Credit Score Adjustors:  (lowest middle score)"
                  primary_key = "LoanType"
                  @loan_amount[primary_key] = {}
                end
                # Loan Size Adjustors
                if r >= 53 && r <= 55 && cc == 4
                  secondary_key = get_value value.split("Loan Amount:").last
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 53 && r <= 55 && cc == 13
                  @adjustment_hash[primary_key][secondary_key] = value
                end
                # Credit Score Adjustors
                if r >= 57 && r <= 58 && cc == 4
                  secondary_key = value
                  @loan_amount[primary_key][secondary_key] = {}
                end
                if r >= 57 && r <= 58 && cc == 13
                  @loan_amount[primary_key][secondary_key] = value
                end
              end
            end
          end
        end
        # adjustment = [@adjustment_hash,@loan_amount,@other_adjustment]
        # make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_united_wholesale_mortgage4892_path(@sheet_obj)
  end

  def non_conf
    @xlsx.sheets.each do |sheet|
      if (sheet == "Non-Conf")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @loan_amount = {}
        @other_adjustment = {}
        primary_key = ''
        secondary_key = ''
        ltv_key = ''
        #program
        (12..38).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3)) && (!row.compact.include?("ELITE PROGRAM IS DESIGNED FOR BORROWERS WITH 700+ CREDIT SCORES"))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 3 # 3 / 7 / 11
              @title = sheet_data.cell(r,cc)
              if @title.present? 
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property
              end
              @block_hash = {}
              key = ''
              (1..11).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
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
        (41..53).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 2)) 
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 5 # 3 / 7 / 11
              @title = sheet_data.cell(r,cc)
              if @title.present? 
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property
              end

              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..11).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
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
        (54..66).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3)) && (!row.compact.include?("ELITE PROGRAM IS DESIGNED FOR BORROWERS WITH 700+ CREDIT SCORES"))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 3 # 3 / 7 / 11
              @title = sheet_data.cell(r,cc)
              if @title.present? 
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property
              end

              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..11).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
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

        # adjustments
        (69..93).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(69)
          if row.compact.count > 1
            (0..14).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Credit Score"
                  primary_key = "LoanType/FICO/CLTV"
                  @adjustment_hash[primary_key] = {}
                end
                if value == "Loan Amount:"
                  primary_key = "LoanType/LoanAmount/CLTV"
                  @loan_amount[primary_key] = {}
                end
                if value == "Cash Out"
                  primary_key = "RefinanceOption/LTV"
                  @other_adjustment[primary_key] = {}
                end
                if value == "Purchase"
                  primary_key = "LoanPurpose"
                  @other_adjustment[primary_key] = {}
                end
                if value == "Escrow Waiver (LTVs >80%; CA only)"
                  primary_key = "Escrow Waiver"
                  @other_adjustment[primary_key] = {}
                end
                # Credit Score
                if r >= 70 && r <= 75 && cc == 5
                  secondary_key = get_value value
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 70 && r <= 75 && cc >= 8 && cc <= 14
                  ltv_key = get_value @ltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = value
                end
                # Loan Amount:
                if r >= 76 && r <= 82 && cc == 5
                  secondary_key = get_value value
                  @loan_amount[primary_key][secondary_key] = {}
                end
                if r >= 76 && r <= 82 && cc >= 8 && cc <= 14
                  ltv_key = get_value @ltv_data[cc-1]
                  @loan_amount[primary_key][secondary_key][ltv_key] = {}
                  @loan_amount[primary_key][secondary_key][ltv_key] = value
                end
                # Other Adjustment
                if r == 83 && cc >= 8 && cc <= 14
                  ltv_key = get_value @ltv_data[cc-1]
                  @other_adjustment[primary_key][ltv_key] = {}
                  @other_adjustment[primary_key][ltv_key] = value
                end
                if r >= 84 && r <= 88 && cc == 3
                  primary_key = value
                  @other_adjustment[primary_key] = {}
                end
                if r >= 84 && r <= 88 && cc >= 8 && cc <= 14
                  ltv_key = get_value @ltv_data[cc-1]
                  @other_adjustment[primary_key][ltv_key] = {}
                  @other_adjustment[primary_key][ltv_key] = value
                end
                if r == 89 && cc >= 8 && cc <= 14
                  ltv_key = get_value @ltv_data[cc-1]
                  @other_adjustment[primary_key][ltv_key] = {}
                  @other_adjustment[primary_key][ltv_key] = value
                end
                if r == 90 && cc >= 8 && cc <= 14
                  ltv_key = get_value @ltv_data[cc-1]
                  @other_adjustment[primary_key][ltv_key] = {}
                  @other_adjustment[primary_key][ltv_key] = value
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@loan_amount,@other_adjustment]
        make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_united_wholesale_mortgage4892_path(@sheet_obj)
  end

  def programs
    @programs = @sheet_obj.programs
  end
  
  def single_program
  end

  def get_program
    @program = Program.find(params[:id])
  end

	private
    def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end

    def get_value value1
      if value1.present?
        if value1.include?("CLTV  <")
          value1 = "0"+value1.split("CLTV").last
        elsif value1.include?("<=") || value1.include?("<")
          value1 = "0"+value1
        elsif value1.include?("CLTV")
          value1 = value1.split("CLTV ").last.first(9)
        else
          value1
        end
      end
    end

    def make_adjust(block_hash, sheet)
      block_hash.each do |hash|
        Adjustment.create(data: hash,sheet_name: sheet)
      end
    end

    def read_sheet
      file = File.join(Rails.root,  'OB_United_Wholesale_Mortgage4892.xls')
      @xlsx = Roo::Spreadsheet.open(file)
    end

    def program_property
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
      if @program.program_name.include?("High Bal") || @program.program_name.include?("High Balance")
        jumbo_high_balance = true
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

      # Arm Advanced
      if @program.program_name.include?("2-2-5 ")
        arm_advanced = "2-2-5"
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
      @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline, jumbo_high_balance: jumbo_high_balance, arm_basic: arm_basic, arm_advanced: arm_advanced)
    end

    # create programs
    def make_program start_range, end_range, sheet_data, row_count, column_count, num1, num2
      (start_range..end_range).each do |r|
        row = sheet_data.row(r)
        if ((row.compact.count > 1) && (row.compact.count <= 4)) && (!row.compact.include?("GOVERNMENT PRICE ADJUSTMENTS"))
          rr = r + 1 
          max_column_section = row.compact.count - 1
          (0..max_column_section).each do |max_column|
            cc = num1 + max_column*num2
            @title = sheet_data.cell(r,cc)
            if @title.present?
              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              @programs_ids << @program.id
              # term
              if @title.scan(/\d+/).count == 1
                term = @title.scan(/\d+/)[0]
              else  
                term = (@title.scan(/\d+/)[0]+ @title.scan(/\d+/)[1]).to_i
              end
              # High Balance
              if @title.include?("HIGH BAL") || @title.include?("HIGH BALANCE")
                loan_size = "High Balance"
                jumbo_high_balance = true
              end
              # Loan-Type
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

              # Streamline Vha, Fha, Usda
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
              end

              # Update Program
              @program.update(term: term, loan_size: loan_size, jumbo_high_balance: jumbo_high_balance,loan_type: loan_type,fha: fha, va: va, usda: usda, streamline: streamline, full_doc: full_doc)
              # Base rate
              @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..row_count).each do |max_row|
                @data = []
                (0..column_count).each_with_index do |index, c_i|
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
      end
    end
end
