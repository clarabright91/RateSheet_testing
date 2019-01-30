class ObHomePointFinancialWholesale11098Controller < ApplicationController
  before_action :get_sheet, only: [:programs, :conforming_standard]
  before_action :get_program, only: [:single_program, :program_property]

  def index
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "Conforming Standard")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "Home Point Financial Corporation"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue  
      # the required headers are not all present
    end
  end

  def conforming_standard
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Conforming Standard")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}
  
        (9..101).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 2))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 8*max_column + 1

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
              if @program.term.present? 
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    elsif (c_i == 1)
                      @block_hash[main_key][key][21] = value
                    elsif (c_i == 2)
                      @block_hash[main_key][key][30] = value
                    elsif (c_i == 3)
                      @block_hash[main_key][key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # Adjustments
        # (67..87).each do |r|
        #   row = sheet_data.row(r)
        #   @key_data = sheet_data.row(40)
        #   if (row.compact.count >= 1)
        #     (0..7).each do |max_column|
        #       cc = max_column
        #       value = sheet_data.cell(r,cc)
        #       if value.present?
        #         if value == "GOVERNMENT ADJUSTMENTS"
        #           first_key = "GovermentAdjustments"
        #           @data_hash[first_key] = {}
        #         end

        #         if value == "FICO, LOAN AMOUNT & PROPERTY TYPE ADJUSTMENTS"
        #           second_key = "FicoLoanAmont"
        #           @data_hash[first_key][second_key] = {}
        #         end

        #         if r >= 70 && r <= 87 && cc == 1
        #           value = get_value value
        #           ccc = cc + 6
        #           c_val = sheet_data.cell(r,ccc)
        #           @data_hash[first_key][second_key][value] = c_val
        #         end

        #       end
        #     end

        #     (10..16).each do |max_column|
        #       cc = max_column
        #       value = sheet_data.cell(r,cc)
        #       if value.present?
        #         if value == "MISCELLANEOUS"
        #           first_key1 = "GovermentAdjustments"
        #           second_key1 = "Miscellaneous"
        #           @misc_hash[first_key1] = {}
        #           @misc_hash[first_key1][second_key1] = {}
        #         end

        #         if r >= 70 && r <= 77 && cc == 10
        #           value1 = get_key value
        #           ccc = cc + 6
        #           c_val = sheet_data.cell(r,ccc)
        #           @misc_hash[first_key1][second_key1][value] = c_val
        #         end

        #         if value == "STATE ADJUSTMENTS"
        #           first_key2 = "GovermentAdjustments"
        #           second_key2 = "StateAdjustments"
        #           @state_hash[first_key2] = {}
        #           @state_hash[first_key2][second_key2] = {}
        #         end

        #         if r >= 80 && r <= 87 && cc == 11
        #           adj_key = value.split(', ')
        #           adj_key.each do |f_key|
        #             key_val = f_key
        #             ccc = cc + 5
        #             k_val = sheet_data.cell(r,ccc)
        #             @state_hash[first_key2][second_key2][key_val] = k_val
        #           end
        #         end

        #       end
        #     end
        #   end
        # end
        # Adjustment.create(data: @data_hash, sheet_name: sheet)
        # Adjustment.create(data: @misc_hash, sheet_name: sheet)
        # Adjustment.create(data: @state_hash, sheet_name: sheet)
      end
    end
    # redirect_to programs_import_file_path(@bank)
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end
  

  def get_value value1
    if value1.present?
      if value1.include?("FICO <")
        value1 = "0"+value1.split("FICO").last
      elsif value1.include?("<")
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



