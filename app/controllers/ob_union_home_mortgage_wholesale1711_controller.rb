class ObUnionHomeMortgageWholesale1711Controller < ApplicationController
  before_action :get_sheet, only: [:programs, :conventional]
	before_action :get_program, only: [:single_program, :program_property]

	def index
    file = File.join(Rails.root,'OB_Union_Home_Mortgage_Wholesale1711.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "Conventional")
          # headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "Union Home Mortgage Wholesale"
          
          @bank = Bank.find_or_create_by(name: @name)
        end
        # @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end

  end

  def programs
  	
  end

  def get_sheet
    @sheet_obj = Sheet.find(params[:id])
  end

  def get_program
    @program = Program.find(params[:id])
  end

	def conventional
		file = File.join(Rails.root,  'OB_Union_Home_Mortgage_Wholesale1711.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Conventional")
      	sheet_data = xlsx.sheet(sheet)
      	(13..103).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
              
          end
        end
      	
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
	end

	private
    def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end
    def program_property
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
