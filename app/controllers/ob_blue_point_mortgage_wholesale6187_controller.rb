class ObBluePointMortgageWholesale6187Controller < ApplicationController

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

  def fha_standard_programs
    file = File.join('/home/yuva/Downloads/sheets_of_remote_url/OB_BluePoint_Mortgage_Wholesale6187.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      sheet_data = xlsx.sheet(sheet)
      (104..136).each do |r|
        row = sheet_data.row(r)
        (1..3).each do |col|
          @title = sheet_data.cell(r,col)
        end
      end
    end
  end

  private
  def get_sheets_names
    return ["FHA STANDARD PROGRAMS", "FHA STREAMLINE PROGRAMS", "VA STANDARD PROGRAMS", "VA STREAMLINE PROGRAMS", "CONVENTIONAL FIXED PROGRAMS", "CONVENTIONAL ARM PROGRAMS", "CONVENTIONAL PRICE ADJUSTMENTS", "FREDDIE MAC PROGRAMS", "FREDDIE MAC PRICE ADJUSTMENTS", "Core Jumbo - Minimum Loan Amount $1.00 above Agency Limit", "Choice Advantage Plus", "Choice Advantage", "Choice Alternative", "Choice Ascent", "Choice Investor", "Leverage - Prime", "Leverage - Lite", "Leverage - Investor", "Leverage - Investor DSCR", "Pivot Prime Jumbo", "Pivot Core / Plus"]
  end
end
