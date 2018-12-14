class ImportFilesController < ApplicationController
  before_action :get_bank, only: [:import_government_sheet, :programs, :import_freddie_fixed_rate, :import_conforming_fixed_rate, :home_possible, :conforming_arms]

  require 'roo'
  require 'roo-xls'

  def index
    xlsx = Roo::Spreadsheet.open("/home/yuva/Desktop/Projects/Pure-Loan_last/RateSheetExtractor/OB_New_Penn_Financial_Wholesale5806.xls")
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
      end
    rescue
      # the required headers are not all present
    end
  end


  def import_government_sheet
    xlsx = Roo::Spreadsheet.open("/home/yuva/Desktop/Projects/Pure-Loan_last/RateSheetExtractor/OB_New_Penn_Financial_Wholesale5806.xls")
    xlsx.sheets.each do |sheet|
      if (sheet == "Government")
        sheet_data = xlsx.sheet(sheet)
        (1..95).each do |r|
          row = sheet_data.row(r)

          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)

              @title = sheet_data.cell(r,cc)
              program_heading = @title.split
              @term = program_heading[1]
              @interest_type = program_heading[3]

              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0)
              @block_hash = {}
              key = ''
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[key] = {}
                  else
                    # first_row[c_i]
                    @block_hash[key][15*c_i] = value
                  end
                  @data << value
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end

              @program.update(interest_points: @block_hash)
            end
          end
        end

        xlsx.sheet(sheet).each_with_index do |sheet_row, index|
          index = index+ 1
          if sheet_row.include?("Loan Level Price Adjustments")
            (index..xlsx.sheet(sheet).last_row).drop(1).each do |row_no|
              adj_row = xlsx.sheet(sheet).row(row_no)
              if xlsx.sheet(sheet).row(row_no).compact.length > 0
                rr = row_no # (r == 8) / (r == 36) / (r == 56)
                max_column_section = adj_row.compact.count - 1
                (0..max_column_section).each do |max_column|
                  cc = 3 + max_column*8 # (3 / 9 / 15)
                  @block_hash = {}
                  key = ''
                  (0..50).each do |max_row|
                    @data = []
                    (0..7).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = xlsx.sheet(sheet).cell(rrr,ccc)
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      elsif (index == 2)
                        @block_hash[key][value.split[0]] = {}
                        # first_row[c_i]
                        # {"Credit Score"=> {0 => 4.0, 580 => 2.00, 600 => 1.250},

                        # {"Credit Score"=>{15=>nil, 30=>"< 580", 45=>nil, 60=>nil, 75=>nil, 90=>4.0, 105=>nil}}
                        # debugger

                        debugger
                        # @block_hash[key][(value.include?("<") ? value.split[0] : nil ] = value if !(value.include?("<"))
                      elsif (index == 6)
                        @block_hash[key][value.split[0]] = value
                      end
                      @data << value
                    end

                    if @data.compact.reject { |c| c.blank? }.length == 0
                      break # terminate the loop
                    end
                  end
                  @program.update(interest_points: @block_hash)
                end
              end
            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def import_freddie_fixed_rate
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Freddie Fixed Rate")
        sheet_data = xlsx.sheet(sheet)

        (1..118).each do |r|
          row = sheet_data.row(r)

          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)

              @title = sheet_data.cell(r,cc)
              program_heading = @title.split
              if @title.scan(/\d+/)[0] == "10"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "15"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "20"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "25"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "30"
                @term = @title.scan(/\d+/)[0]
              end
              if program_heading[3] == "Fixed"
                @interest_type = 0
              end
              if program_heading[0] + program_heading[1] == "FreddieMac"
                @conforming = true
                @freddie_mac = true
              end
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac)
              @block_hash = {}
              key = ''
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  
                  if (c_i == 0)
                    key = value
                    @block_hash[key] = {}
                  else
                    # first_row[c_i]
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)
                    else
                      @block_hash[key][15*c_i] = value
                    end
                  end
                  
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash)
            end
          end
        end
        # (121..178).each do |r|
        #   row = sheet_data.row(r)
        #   if row.compact.count > 1
        #     rr = r+1
        #     max_column_section = row.count-1
        #     (0..max_column_section).each do |max_column|
        #       cc = max_column
        #       @block_adjustment_hash = {}
        #       # (0..50).each do |max_row|
        #         @adjustment_data = []
        #         (0..19).each_with_index do |index, c_i|
        #           rrr = rr
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)

        #           if (c_i == 0)
        #             key = value
        #             @block_adjustment_hash[key] = {}
        #           else
        #             @block_adjustment_hash["key"] = "All Fixed Conforming\n(does not apply to terms <=15yrs)"
        #             debugger
        #             if row.compact.include?("< 620")
        #               @adjustment_headers = row.compact
        #             end
        #             @block_adjustment_hash[key][@adjustment_headers[0]] = value
        #           end
        #           # debugger
        #           @adjustment_data << value
        #         end
        #       # end
        #     end
        #   end
        # end
        # (123..178).each do |r|
        #   row = sheet_data.row(r)
        #   if row.compact.count > 1
        #     max_column_section = row.count-1
        #     @block_adjustment_hash = {}
        #     @adjustment_data = []
        #     (0..max_column_section).each_with_index do |max_column|
        #       value = sheet_data.cell(r,max_column)

        #       if (value == "All Fixed Conforming\n(does not apply to terms <=15yrs)")
        #         key = value
        #         @block_adjustment_hash[key] = {}
        #       elsif value.present?
        #         debugger
        #         @block_adjustment_hash[key][max_column] = value
        #       end
        #       @adjustment_data << value
        #       debugger
        #     end
        #   end
        # end
        (122..178).each do |r|
          row = sheet_data.row(r)
          if r == 122
            @adjustment_headers = row.compact
          end
          if row.compact.count > 1
            max_column_section = row.count-1
            @block_adjustment_hash = {}
            @adjustment_data = []
            key = ''
            first_column = ''
            column = 3
            (0..max_column_section).each_with_index do |index, max_column|
              cc = column + max_column
              value = sheet_data.cell(r,cc)
              # debugger if max_column == 7
              if value.present? && value != "LTV"
                if (value == "All Fixed Conforming\n(does not apply to terms <=15yrs)")
                  key = value
                  @block_adjustment_hash[key] = {}
                elsif (value.class ==String) && (value.include?("<="))
                  first_column = ((value.include?("<=") || value.include?("<")) ? "0" : value.split.first)
                  @block_adjustment_hash[key][first_column] = first_column
                elsif value.is_a?(Numeric)
                  @adjustment_headers.each do |header|
                    @i_hash = {}
                    new_val = {header => value}
                    @i_hash.merge!(new_val)
                    @block_adjustment_hash[key][first_column] = {} #{header => value}
                    debugger
                  end
                  @block_adjustment_hash[key][first_column].merge!(@i_hash)
                end
                @adjustment_data << value
              end
            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def import_conforming_fixed_rate
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Conforming Fixed Rate")
        sheet_data = xlsx.sheet(sheet)

        (1..118).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)

              @title = sheet_data.cell(r,cc)
              program_heading = @title.split
              if @title.scan(/\d+/)[0] == "10"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "15"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "20"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "25"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "30"
                @term = @title.scan(/\d+/)[0]
              end
              if program_heading[3] == "Fixed"
                @interest_type = 0
              end
              if program_heading[0] + program_heading[1] == "FreddieMac"
                @conforming = true
                @freddie_mac = true
              end
              if program_heading[0]+program_heading[1] =="FannieMae"
                @fannie_mae = true
              end
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae)
              @block_hash = {}
              key = ''
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[key] = {}
                  else
                    # first_row[c_i]
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash)
            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def home_possible
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Home Possible")
        sheet_data = xlsx.sheet(sheet)

        (1..76).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)

              @title = sheet_data.cell(r,cc)
              program_heading = @title.split
              if @title.scan(/\d+/)[0] == "10"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "15"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "20"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "25"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "30"
                @term = @title.scan(/\d+/)[0]
              end
              @title.include?("30")
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae)
              @block_hash = {}
              key = ''
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[key] = {}
                  else
                    # first_row[c_i]
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash)
            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def conforming_arms
    file = File.join(Rails.root,  'OB_New_Penn_Financial_Wholesale5806.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Conforming ARMs")
        sheet_data = xlsx.sheet(sheet)

        (1..47).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)

              @title = sheet_data.cell(r,cc)
              program_heading = @title.split
              if @title.scan(/\d+/)[0] == "10"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "15"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "20"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "25"
                @term = @title.scan(/\d+/)[0]
              elsif @title.scan(/\d+/)[0] == "30"
                @term = @title.scan(/\d+/)[0]
              end
              if @title.include?("Fixed")
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              end
              if @title.include?("Freddie Mac")
                @freddie_mac = true
              end
              if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
                @conforming = true
              end
              if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
                @fannie_mae = true
              end
              if @title.include?("High Balance")
                @jumbo_high_balance = true
              end
              
              @program = @bank.programs.find_or_create_by(title: @title)
              @program.update(term: @term,interest_type: @interest_type,loan_type: 0,conforming: @conforming,freddie_mac: @freddie_mac, fannie_mae: @fannie_mae, jumbo_high_balance: @jumbo_high_balance)
              @block_hash = {}
              key = ''
              (0..50).each do |max_row|
                @data = []
                (0..4).each_with_index do |index, c_i|
                  rrr = rr + max_row
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if (c_i == 0)
                    key = value
                    @block_hash[key] = {}
                  else
                    # first_row[c_i]
                    if value.class == Float
                      @block_hash[key][15*c_i] = value.round(3)  
                    else
                      @block_hash[key][15*c_i] = value
                    end
                  end
                  @data << value
                end

                if @data.compact.length == 0
                  break # terminate the loop
                end
              end
              @program.update(interest_points: @block_hash)
            end
          end
        end
      end
    end
    redirect_to programs_import_file_path(@bank)
  end

  def programs
    @programs = @bank.programs
  end

  private

  def get_bank
    @bank = Bank.find(params[:id])
  end
end