class ObBluePointMortgageWholesale6187Controller < ApplicationController
  include SheetOperation

  def index
    sub_sheet_names = get_sheets_names
    file = File.join('/home/yuva/Downloads/sheetsfor3rdmilestone/OB_BluePoint_Mortgage_Wholesale6187.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "Blue Point")
          @name = "BluePoint Mortgage"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
        sub_sheet_names.each do |sub_sheet|
          @sub_sheet = @sheet.sub_sheets.create(name: sub_sheet)
        end
      end
    rescue
    end
  end

  def fha_streamline_programs
  end
end
