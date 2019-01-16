class ObCardinalFinancialWholesale10742Controller < ApplicationController
	before_action :get_sheet, only: [:programs, :ak]
	before_action :get_program, only: [:single_program]
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
        (71..1260).each do |r|
        	row = sheet_data.row(r)
        	if row.compact.count > 1 && row.compact.count <= 3
        		max_column_section = row.compact.count - 1
        		(0..max_column_section).each do |max_column|
        			cc = max_column + 1
        			if value.present?
	        			if cc == 2 || cc == 13 || cc == 25 || cc == 35
	        				@title = sheet_data.cell(r,cc)
	        			end
	        		end
        		end
        		debugger
        	end
        end
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
end
