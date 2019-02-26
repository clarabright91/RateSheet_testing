class ObAmericanFinancialResourcesWholesale5513Controller < ApplicationController
  before_action :get_sheet, only: [:programs, :gnma, :gnma_hb, :fnma, :fhlmc, :hp, :jumbo]
  before_action :read_sheet, only: [:gnma, :gnma_hb, :fnma, :fhlmc, :hp, :jumbo]
  before_action :get_program, only: [:single_program, :program_property]

  def index
    file = File.join(Rails.root, 'OB_American_Financial_Resources_Wholesale5513.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "GNMA")
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

  def gnma
    @xlsx.sheets.each do |sheet|
      if (sheet == "GNMA")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        start_range = 10
        end_range = 30
        row_count = 20
        column_count = 3
        num1 = 4
        num2 = 1
        make_program start_range, end_range, sheet_data, row_count, column_count, num1, num2
      end
    end
    redirect_to programs_ob_american_financial_resources_wholesale5513_path(@sheet_obj)
  end

  def gnma_hb
    @xlsx.sheets.each do |sheet|
      if (sheet == "GNMA HB")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        start_range = 10
        end_range = 30
        row_count = 20
        column_count = 3
        num1 = 4
        num2 = 1
        make_program start_range, end_range, sheet_data, row_count, column_count, num1, num2
      end
    end
    redirect_to programs_ob_american_financial_resources_wholesale5513_path(@sheet_obj)
  end

  def fnma
    @xlsx.sheets.each do |sheet|
      if (sheet == "FNMA")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        start_range = 10
        end_range = 72
        row_count = 20
        column_count = 3
        num1 = 4
        num2 = 1
        make_program start_range, end_range, sheet_data, row_count, column_count, num1, num2
      end
    end
    redirect_to programs_ob_american_financial_resources_wholesale5513_path(@sheet_obj)
  end

  def fhlmc
    @xlsx.sheets.each do |sheet|
      if (sheet == "FHLMC")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        start_range = 10
        end_range = 72
        row_count = 20
        column_count = 3
        num1 = 4
        num2 = 1
        make_program start_range, end_range, sheet_data, row_count, column_count, num1, num2
      end
    end
    redirect_to programs_ob_american_financial_resources_wholesale5513_path(@sheet_obj)
  end

  def hp
    @xlsx.sheets.each do |sheet|
      if (sheet == "HP")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        start_range = 10
        end_range = 30
        row_count = 20
        column_count = 3
        num1 = 4
        num2 = 1
        make_program start_range, end_range, sheet_data, row_count, column_count, num1, num2
      end
    end
    redirect_to programs_ob_american_financial_resources_wholesale5513_path(@sheet_obj)
  end

  def jumbo
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        start_range = 11
        end_range = 27
        row_count = 20
        column_count = 3
        num1 = 5
        num2 = 1
        make_program start_range, end_range, sheet_data, row_count, column_count, num1, num2
      end
    end
    redirect_to programs_ob_american_financial_resources_wholesale5513_path(@sheet_obj)
  end

  private
    def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end

    def read_sheet
      file = File.join(Rails.root,  'OB_American_Financial_Resources_Wholesale5513.xls')
      @xlsx = Roo::Spreadsheet.open(file)
    end

    # def program_property
    #   if @program.program_name.include?("30") || @program.program_name.include?("30/25 Year")
    #     term = 30
    #   elsif @program.program_name.include?("20") || @program.program_name.include?("20 Year")
    #     term = 20
    #   elsif @program.program_name.include?("15")
    #     term = 15
    #   elsif @program.program_name.include?("10 Year")
    #     term = 10
    #   elsif @program.program_name.include?("5/1")
    #     term = 5
    #   else
    #     term = nil
    #   end

    #   # Loan-Type
    #   if @program.program_name.include?("Fixed") || @program.program_name.include?("FIXED")
    #     loan_type = "Fixed"
    #   elsif @program.program_name.include?("ARM")
    #     loan_type = "ARM"
    #   elsif @program.program_name.include?("Floating")
    #     loan_type = "Floating"
    #   elsif @program.program_name.include?("Variable")
    #     loan_type = "Variable"
    #   else
    #     loan_type = nil
    #   end

    #   # Streamline Vha, Fha, Usda
    #   fha = false
    #   va = false
    #   usda = false
    #   streamline = false
    #   full_doc = false
    #   if @program.program_name.include?("FHA")
    #     streamline = true
    #     fha = true
    #     full_doc = true
    #   elsif @program.program_name.include?("VA")
    #     streamline = true
    #     va = true
    #     full_doc = true
    #   elsif @program.program_name.include?("USDA")
    #     streamline = true
    #     usda = true
    #     full_doc = true
    #   end

    #   # High Balance
    #   jumbo_high_balance = false
    #   if @program.program_name.include?("High Bal") || @program.program_name.include?("High Balance")
    #     jumbo_high_balance = true
    #   end

    #   # Arm Basic
    #   if @program.program_name.include?("3/1") || @program.program_name.include?("3 / 1")
    #     arm_basic = 3
    #   elsif @program.program_name.include?("5/1") || @program.program_name.include?("5 / 1")
    #     arm_basic = 5
    #   elsif @program.program_name.include?("7/1") || @program.program_name.include?("7 / 1")
    #     arm_basic = 7
    #   elsif @program.program_name.include?("10/1") || @program.program_name.include?("10 / 1")
    #     arm_basic = 10
    #   end

    #   # Arm Advanced
    #   if @program.program_name.include?("2-2-5 ")
    #     arm_advanced = "2-2-5"
    #   end
    #   # Loan Limit Type
    #   if @program.program_name.include?("Non-Conforming")
    #     @program.loan_limit_type << "Non-Conforming"
    #   end
    #   if @program.program_name.include?("Conforming")
    #     @program.loan_limit_type << "Conforming"
    #   end
    #   if @program.program_name.include?("Jumbo")
    #     @program.loan_limit_type << "Jumbo"
    #   end
    #   if @program.program_name.include?("High Balance")
    #     @program.loan_limit_type << "High Balance"
    #   end
    #   @program.save
    #   @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline, jumbo_high_balance: jumbo_high_balance, arm_basic: arm_basic, arm_advanced: arm_advanced)
    # end

    def program_property
      if @program.program_name.include?("30") || @program.program_name.include?("30/25 Year")
        term = 30
      elsif @program.program_name.include?("20")
        term = 20
      elsif @program.program_name.include?("15")
        term = 15
      elsif @program.program_name.include?("10 Year")
        term = 10
      else
        term = nil
      end

      # Loan-Type
      if @program.program_name.include?("Fixed") || @program.program_name.include?("FIXED")
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
        if ((row.compact.count > 1) && (row.compact.count <= 6)) && (!row.compact.include?("GOVERNMENT PRICE ADJUSTMENTS"))
          rr = r + 1
          max_column_section = row.compact.count - 1
          (0..max_column_section).each do |max_column|
            cc = num1 * max_column + num2
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
