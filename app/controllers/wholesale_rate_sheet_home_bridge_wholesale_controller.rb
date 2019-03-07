class WholesaleRateSheetHomeBridgeWholesaleController < ApplicationController
	before_action :get_sheet, only: [:rate_sheet, :programs, :program_property]
  before_action :get_program, only: [:single_program, :program_property]
  before_action :read_sheet, only: [:index,:rate_sheet,:programs]

  def index
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "Rate Sheet")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "HomeBridge Wholesale"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def rate_sheet
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "Rate Sheet")
      	debugger
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

  def read_sheet
    file = File.join(Rails.root,  'Wholesale Rate Sheet _HomeBridge Wholesale_.xls')
    @xlsx = Roo::Spreadsheet.open(file)
  end

  def get_program
    @program = Program.find(params[:id])
  end
end