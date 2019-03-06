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
    file = File.join('/home/yuva/Downloads/sheetsfor3rdmilestone/OB_BluePoint_Mortgage_Wholesale6187.xls')
    xlsx = Roo::Spreadsheet.open(file)
    @sheet_data = {}
    xlsx.sheets.each do |sheet|
      sheet_data = xlsx.sheet(sheet)
      main_key = nil
      (101..136).each do |r|
        row = sheet_data.row(r)

        # find sub sheet according to sheet name
        if(get_sheets_names.include?(row.first))
          @sub_sheet = SubSheet.find_by_name(row.first)
        end

        if (r > 103) && @sub_sheet.present?
          [2, 6, 10].each do |c|
            @title = sheet_data.cell(r,c)
            (c..c+3).each do |col|
              value = sheet_data.cell(r,col)
              if(r > 106) && (c.eql?(col))
                unless @sheet_data.has_key?(@title)
                  @sheet_data[@title] = {}
                  main_key = @title
                end
              elsif r > 106 && col > c && col <= c + 3
                secondary_key = sheet_data.cell(106,col)
                @sheet_data[main_key][secondary_key.to_i] = value
              end

              if r > 106 && c+3 == col
              end
            end
          end
        end
      end
    end
  end

  private
  def get_sheets_names
    return ["FHA STANDARD PROGRAMS", "FHA STREAMLINE PROGRAMS", "VA STANDARD PROGRAMS", "VA STREAMLINE PROGRAMS", "CONVENTIONAL FIXED PROGRAMS", "CONVENTIONAL ARM PROGRAMS", "CONVENTIONAL PRICE ADJUSTMENTS", "FREDDIE MAC PROGRAMS", "FREDDIE MAC PRICE ADJUSTMENTS", "Core Jumbo - Minimum Loan Amount $1.00 above Agency Limit", "Choice Advantage Plus", "Choice Advantage", "Choice Alternative", "Choice Ascent", "Choice Investor", "Leverage - Prime", "Leverage - Lite", "Leverage - Investor", "Leverage - Investor DSCR", "Pivot Prime Jumbo", "Pivot Core / Plus"]
  end
end
