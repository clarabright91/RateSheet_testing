class ImportFilesController < ApplicationController
  def index
    require 'roo'
    require 'roo-xls'
    xlsx = Roo::Spreadsheet.open("/home/richa/richa/Kevin-Project/Pure-Loan/OB_New_Penn_Financial_Wholesale5806.xls")
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "Cover Zone 1")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          xlsx.sheet(sheet).each_with_index do |row, index|
            current_row = index+1
            if row.include?("Mortgagee Clause (Wholesale)")
              address_index = row.find_index("Mortgagee Clause (Wholesale)")
              @address_a = []
              (1..3).each do |n|
                @address_a << xlsx.sheet(sheet).row(current_row+n)[address_index]
                if n == 3
                  @zip = xlsx.sheet(sheet).row(current_row+n)[address_index].split.last
                  @state_code = xlsx.sheet(sheet).row(current_row+n)[address_index].split[2]
                end
              end
            end
            if (row.include?("Phone") && row.include?("General Contacts"))
              phone_index = row.find_index(headers[0])
              general_contacts_index = row.find_index(headers[1])
              c_row = xlsx.sheet(sheet).row(current_row+1)
              @name = c_row[general_contacts_index]
              @phone = c_row[phone_index]
            end
          end
          @bank = Bank.find_or_create_by(name: @name)
          @bank.update(phone: @phone, address1: @address_a.join, state_code: @state_code, zip: @zip)
        end
        if (sheet == "Government")
          sheet_data = xlsx.sheet(sheet)

          (1..95).each do |r|
            row = sheet_data.row(r)

            if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
              # (0..row.compact.count).each do |header_col|
              #   term = row.compact[header_col]
              #   debugger
              # end
              # r == 7 / 35 / 55
              rr = r + 1 # (r == 8) / (r == 36) / (r == 56)

              max_column_section = row.compact.count - 1
              (0..max_column_section).each do |max_column|
                cc = 3 + max_column*6 # (3 / 9 / 15) 3/8/13

                @title = sheet_data.cell(r,cc)
                program_heading = @title.split
                @term = program_heading[1]
                @interest_type = program_heading[3]

                @program = @bank.programs.find_or_create_by(title: @title)
                @program.update(term: @term,interest_type: @interest_type,loan_type: 0)
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    # i_hash["interest"] = value
                    @data << value
                  end

                  i_hash = Hash.new
                  i_hash[@data[0]] = ["15 Day", "30 Day", "45 Day", "60 Day"].zip(@data.drop(1)).to_h
                  debugger
                  # hash = {}
                  # @data.each { |i| hash[i] = 'free' }

                  if @data.compact.length == 0
                    # terminate the loop
                    break
                  end
                end
              end
            end
          end
        end
      end
    rescue
      # the required headers are not all present
    end
  end

  def new
  end

  def create
  end
end


