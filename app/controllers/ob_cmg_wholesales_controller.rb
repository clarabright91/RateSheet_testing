class ObCmgWholesalesController < ApplicationController
	# before_action :get_sheet, only: [:import_gov_sheet]
  def index
  	file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "AGENCY")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          # xlsx.sheet(sheet).each_with_index do |row, index|
          #   current_row = index+1
          #   if row.include?("Mortgagee Clause (Wholesale)")
          #     address_index = row.find_index("Mortgagee Clause (Wholesale)")
          #     @address_a = []
          #     (1..3).each do |n|
          #       @address_a << xlsx.sheet(sheet).row(current_row+n)[address_index]
          #       if n == 3
          #         @zip = xlsx.sheet(sheet).row(current_row+n)[address_index].split.last
          #         @state_code = xlsx.sheet(sheet).row(current_row+n)[address_index].split[2]
          #       end
          #     end
          #   end
          #   if (row.include?("Phone") && row.include?("General Contacts"))
          #     phone_index = row.find_index(headers[0])
          #     general_contacts_index = row.find_index(headers[1])
          #     c_row = xlsx.sheet(sheet).row(current_row+1)
          #     @name = c_row[general_contacts_index]
          #     @phone = c_row[phone_index]
          #   end
          # end
          @name = "CMG Financial"
          @bank = Bank.find_or_create_by(name: @name)
          # @bank.update(phone: @phone, address1: @address_a.join, state_code: @state_code, zip: @zip)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end
  def import_gov_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "GOV")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (10..60).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)

              # term
              @term = nil
              if @title.include?("30 Year") || @title.include?("30Yr")
                @term = 30
              elsif @title.include?("20 Year")
                @term = 20
              elsif @title.include?("15 Year")
                @term = 15
              end

               	# interest type
              if @title.include?("Fixed")
                @rate_type = 0
              elsif @title.include?("ARM")
                @rate_type = 2
              end

              # streamline
              if @title.include?("FHA") 
                @streamline = true
                @fha = true
                @full_doc = true
              elsif @title.include?("VA")
              	@streamline = true
              	@va = true
              	@full_doc = true
              elsif @title.include?("USDA")
              	@streamline = true
              	@usda = true
              	@full_doc = true
              else
              	@streamline = false
              	@fha = false
              	@va = false
              	@usda = false
              	@full_doc =
              end

              # High Balance
              if @title.include?("High Bal")
              	@jumbo_high_balance = true
              end

              # Program Category
              if @title.include?("3101 & 3125")
              	@program_category = "3101 & 3125"
              elsif @title.include?("3103")
              	@program_category = "3103"
              elsif @title.include?("3102")
              	@program_category = "3102"	
              elsif @title.include?("3101HB & 3125HB")
              	@program_category = "3101HB & 3125HB"
              elsif @title.include?("4101 & 4125")
              	@program_category = "4101 & 4125"	
              elsif @title.include?("4103")
              	@program_category = "4103"
              elsif @title.include?("4102")
              	@program_category = "4102"
              elsif @title.include?("4101HB & 4125HB")
              	@program_category = "4101HB & 4125HB"
              elsif @title.include?("5101")
              	@program_category = "5101"
              elsif @title.include?("3151")
              	@program_category = "3151"	
              elsif @title.include?("4151")
              	@program_category = "4151"
              end

              @program = Program.find_or_create_by(program_name: @title)
              @programs_ids << @program.id
              @program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc)
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    elsif (c_i == 1)
                      @block_hash[key][21] = value
                    elsif (c_i == 2)
                      @block_hash[key][30] = value
                    elsif (c_i == 3)
                      @block_hash[key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end
  def import_agency_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "AGENCY")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (10..87).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)

            	# term
            	@term = nil
            	if @title.present? && @title != "2.250% MARGIN - 2/2/6 CAPS - 1 YR LIBOR" && @title != "2.250% MARGIN - 5/2/5 CAPS - 1 YR LIBOR"
	              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
	                @term = 30
	              elsif @title.include?("20 Year")
	                @term = 20
	              elsif @title.include?("15 Year")
	                @term = 15
	              end
	           
	               	# interest type
	              if @title.include?("Fixed")
	                @rate_type = 0
	              elsif @title.include?("ARM")
	                @rate_type = 2
	              else
	              	@rate_type = nil
	              end

	              # streamline
	              if @title.include?("FHA") 
	                @streamline = true
	                @fha = true
	                @full_doc = true
	              elsif @title.include?("VA")
	              	@streamline = true
	              	@va = true
	              	@full_doc = true
	              elsif @title.include?("USDA")
	              	@streamline = true
	              	@usda = true
	              	@full_doc = true
	              else
	              	@streamline = nil
	              	@full_doc = nil
	              	@fha = nil
	              	@va = nil
	              	@usda = nil
	              end

	              # High Balance
	              if @title.include?("High Bal")
	              	@jumbo_high_balance = true
	              else
	              	@jumbo_high_balance = nil
	              end

	              # rate arm
	              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM") || @title.include?("5/1 LIBOR ARM") || @title.include?("7/1 LIBOR ARM") || @title.include?("10/1 LIBOR ARM")
	                @rate_arm = @title.scan(/\d+/)[0].to_i
	              else
	              	@rate_arm = nil
	              end

	              @program = Program.find_or_create_by(program_name: @title)
	              @programs_ids << @program.id
	              @program.update(term: @term,rate_type: @rate_typerate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
	              # @program.adjustments.destroy_all
	              @block_hash = {}
	              key = ''
	              (1..50).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row -1
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    elsif (c_i == 1)
	                      @block_hash[key][21] = value
	                    elsif (c_i == 2)
	                      @block_hash[key][30] = value
	                    elsif (c_i == 3)
	                      @block_hash[key][45] = value
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	            end
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end
  def import_durp_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "DURP")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (10..53).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)

              # term
              @term = nil
              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
                @term = 30
              elsif @title.include?("20 Year")
                @term = 20
              elsif @title.include?("15 Year")
                @term = 15
              end
              
               	# interest type
              if @title.include?("Fixed")
                @rate_type = 0
              elsif @title.include?("ARM")
                @rate_type = 2
              else
              	@rate_type = nil
              end

              # streamline
              if @title.include?("FHA") 
                @streamline = true
                @fha = true
                @full_doc = true
              elsif @title.include?("VA")
              	@streamline = true
              	@va = true
              	@full_doc = true
              elsif @title.include?("USDA")
              	@streamline = true
              	@usda = true
              	@full_doc = true
              else
              	@streamline = nil
              	@full_doc = nil
              	@fha = nil
              	@va = nil
              	@usda = nil
              end

              # High Balance
              if @title.include?("High Bal")
              	@jumbo_high_balance = true
              else
              	@jumbo_high_balance = nil
              end

              @program = Program.find_or_create_by(program_name: @title)
              @programs_ids << @program.id
              @program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance)
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    elsif (c_i == 1)
                      @block_hash[key][21] = value
                    elsif (c_i == 2)
                      @block_hash[key][30] = value
                    elsif (c_i == 3)
                      @block_hash[key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end
  def import_oa_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "OA")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (10..53).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)

            	# term
            	@term = nil
              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
                @term = 30
              elsif @title.include?("20 Year")
                @term = 20
              elsif @title.include?("15 Year")
                @term = 15
              end
           
               	# interest type
              if @title.include?("Fixed")
                @rate_type = 0
              elsif @title.include?("ARM")
                @rate_type = 2
              else
              	@rate_type = nil
              end

              # streamline
              if @title.include?("FHA") 
                @streamline = true
                @fha = true
                @full_doc = true
              elsif @title.include?("VA")
              	@streamline = true
              	@va = true
              	@full_doc = true
              elsif @title.include?("USDA")
              	@streamline = true
              	@usda = true
              	@full_doc = true
              else
              	@streamline = nil
              	@full_doc = nil
              	@fha = nil
              	@va = nil
              	@usda = nil
              end

              # High Balance
              if @title.include?("High Bal")
              	@jumbo_high_balance = true
              else
              	@jumbo_high_balance = nil
              end
              
              @program = Program.find_or_create_by(program_name: @title)
              @programs_ids << @program.id
              @program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    elsif (c_i == 1)
                      @block_hash[key][21] = value
                    elsif (c_i == 2)
                      @block_hash[key][30] = value
                    elsif (c_i == 3)
                      @block_hash[key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end
  def import_jumbo700_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 700")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        first_key = ''
        @adjustment_hash = {}
        @data_hash = {}
        @sheet = sheet
        @ltv_data = []
        key = ''
        key1 = ''
        key2 = ''
        key3 = ''
        ltv_key = ''
        cltv_key = ''
        c_val = ''
        cc = ''
        value = ''
        state_key = ''
        adj_key = []
        (10..21).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)

            	# term
            	@term = nil
              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
                @term = 30
              elsif @title.include?("20 Year")
                @term = 20
              elsif @title.include?("15 Year")
                @term = 15
              end
           
               	# interest type
              if @title.include?("Fixed")
                @rate_type = 0
              elsif @title.include?("ARM")
                @rate_type = 2
              else
              	@rate_type = nil
              end

              # streamline
              if @title.include?("FHA") 
                @streamline = true
                @fha = true
                @full_doc = true
              elsif @title.include?("VA")
              	@streamline = true
              	@va = true
              	@full_doc = true
              elsif @title.include?("USDA")
              	@streamline = true
              	@usda = true
              	@full_doc = true
              else
              	@streamline = nil
              	@full_doc = nil
              	@fha = nil
              	@va = nil
              	@usda = nil
              end

              # High Balance
              if @title.include?("High Bal")
              	@jumbo_high_balance = true
              else
              	@jumbo_high_balance = nil
              end
              
              @program = Program.find_or_create_by(program_name: @title)
              @programs_ids << @program.id
              @program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    elsif (c_i == 1)
                      @block_hash[key][21] = value
                    elsif (c_i == 2)
                      @block_hash[key][30] = value
                    elsif (c_i == 3)
                      @block_hash[key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end

        # adjustments
        (23..47).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(25)
          if (row.compact.count > 1)
            (0..9).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "ELITE JUMBO 700 SERIES ADJUSTMENTS"
                  first_key = "LoanAmount/FICO/LTV"
                  key = first_key
                  @adjustment_hash[key] = {}
                end
                if value == "Loan Amount <= $1,000,000"
                  key1 = "0"
                  @adjustment_hash[key][key1] = {}
                end
                if value == "Loan Amount > $1,000,000"
                  key1 = "$1,000,000"
                  @adjustment_hash[key][key1] = {}
                end
                if r >= 27 && r <= 33 && cc == 1
                  ltv_key = get_value value
                  @adjustment_hash[key][key1][ltv_key] = {}
                end
                if r >= 27 && r <= 33 && cc > 4 && cc <= 9
                  cltv_key = get_value @ltv_data[cc-1]
                  @adjustment_hash[key][key1][ltv_key][cltv_key] = value
                end
                if r >= 35 && r <= 47 && cc == 1
                  ltv_key = get_value value
                  @adjustment_hash[key][key1][ltv_key] = {}
                end
                if r >= 35 && r <= 47 && cc >= 4 && cc <= 9
                  cltv_key = get_value @ltv_data[cc-1]
                  @adjustment_hash[key][key1][ltv_key][cltv_key] = value
                end
              end
            end

            #For STATE ADJUSTMENTS
            (12..16).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "STATE ADJUSTMENTS"
                  state_key = "StateAdjustments"
                  @adjustment_hash[state_key] = {}
                end
                if r >= 24 && r < 28 && cc == 12
                  adj_key = value.split(', ')
                  adj_key.each do |f_key|
                    key3 = f_key
                    ccc = cc + 4
                    c_val = sheet_data.cell(r,ccc)
                    @adjustment_hash[state_key][key3] = c_val
                  end
                end
              end
            end

            #For MISCELLANEOUS
            (12..16).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "MISCELLANEOUS"
                  state_key = "Miscellaneous"
                  @adjustment_hash[state_key] = {}
                end
                if r == 31 && cc == 12
                  ccc = cc + 4
                  c_val = sheet_data.cell(r,ccc)
                  @adjustment_hash[state_key][value] = c_val
                end
              end
            end
      		end
    		end
    		Adjustment.create(data: @adjustment_hash, sheet_name: sheet)
   		end
   	end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end

  def import_jumbo6200_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 6200")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (10..34).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)

            	# term
            	@term = nil
              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
                @term = 30
              elsif @title.include?("20 Year")
                @term = 20
              elsif @title.include?("15 Year")
                @term = 15
              end
           
               	# interest type
              if @title.include?("Fixed")
                @rate_type = 0
              elsif @title.include?("ARM")
                @rate_type = 2
              else
              	@rate_type = nil
              end

              # streamline
              if @title.include?("FHA") 
                @streamline = true
                @fha = true
                @full_doc = true
              elsif @title.include?("VA")
              	@streamline = true
              	@va = true
              	@full_doc = true
              elsif @title.include?("USDA")
              	@streamline = true
              	@usda = true
              	@full_doc = true
              else
              	@streamline = nil
              	@full_doc = nil
              	@fha = nil
              	@va = nil
              	@usda = nil
              end

              # High Balance
              if @title.include?("High Bal")
              	@jumbo_high_balance = true
              else
              	@jumbo_high_balance = nil
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
                @rate_arm = @title.scan(/\d+/)[0].to_i
              else
              	@rate_arm = nil
              end
              
              @program = Program.find_or_create_by(program_name: @title)
              @programs_ids << @program.id
              @program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    elsif (c_i == 1)
                      @block_hash[key][21] = value
                    elsif (c_i == 2)
                      @block_hash[key][30] = value
                    elsif (c_i == 3)
                      @block_hash[key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end
  def import_jumbo7200_6700_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 7200 & 6700")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (10..22).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
          		if @title.present? && @title.include?("30 Year Fixed - 7230")
	            	# term
	            	@term = nil
	              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
	                @term = 30
	              elsif @title.include?("20 Year")
	                @term = 20
	              elsif @title.include?("15 Year")
	                @term = 15
	              end
	           
	               	# interest type
	              if @title.include?("Fixed")
	                @rate_type = 0
	              elsif @title.include?("ARM")
	                @rate_type = 2
	              else
	              	@rate_type = nil
	              end

	              # streamline
	              if @title.include?("FHA") 
	                @streamline = true
	                @fha = true
	                @full_doc = true
	              elsif @title.include?("VA")
	              	@streamline = true
	              	@va = true
	              	@full_doc = true
	              elsif @title.include?("USDA")
	              	@streamline = true
	              	@usda = true
	              	@full_doc = true
	              else
	              	@streamline = nil
	              	@full_doc = nil
	              	@fha = nil
	              	@va = nil
	              	@usda = nil
	              end

	              # High Balance
	              if @title.include?("High Bal")
	              	@jumbo_high_balance = true
	              else
	              	@jumbo_high_balance = nil
	              end

	              # interest sub type
	              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
	                @rate_arm = @title.scan(/\d+/)[0].to_i
	              else
	              	@rate_arm = nil
	              end
              end
              
              if cc < 5
	              @program = Program.find_or_create_by(program_name: @title)
	              @programs_ids << @program.id
	             	@program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
	            
	              # @program.adjustments.destroy_all
	              @block_hash = {}
	              key = ''
	              (1..50).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row -1
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    elsif (c_i == 1)
	                      @block_hash[key][21] = value
	                    elsif (c_i == 2)
	                      @block_hash[key][30] = value
	                    elsif (c_i == 3)
	                      @block_hash[key][45] = value
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	            end
	            if  @block_hash.values.first.values.first == "21 Day"
              	@block_hash.shift
            	end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        (56..64).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
          		if @title.present? && @title.include?("30 Year Fixed - 7230")
	            	# term
	            	@term = nil
	              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
	                @term = 30
	              elsif @title.include?("20 Year")
	                @term = 20
	              elsif @title.include?("15 Year")
	                @term = 15
	              end
	           
	               	# interest type
	              if @title.include?("Fixed")
	                @rate_type = 0
	              elsif @title.include?("ARM")
	                @rate_type = 2
	              else
	              	@rate_type = nil
	              end

	              # streamline
	              if @title.include?("FHA") 
	                @streamline = true
	                @fha = true
	                @full_doc = true
	              elsif @title.include?("VA")
	              	@streamline = true
	              	@va = true
	              	@full_doc = true
	              elsif @title.include?("USDA")
	              	@streamline = true
	              	@usda = true
	              	@full_doc = true
	              else
	              	@streamline = nil
	              	@full_doc = nil
	              	@fha = nil
	              	@va = nil
	              	@usda = nil
	              end

	              # High Balance
	              if @title.include?("High Bal")
	              	@jumbo_high_balance = true
	              else
	              	@jumbo_high_balance = nil
	              end

	              # interest sub type
	              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
	                @rate_arm = @title.scan(/\d+/)[0].to_i
	              else
	              	@rate_arm = nil
	              end
              end
              
              if cc < 5
	              @program = Program.find_or_create_by(program_name: @title)
	              @programs_ids << @program.id
	             	@program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
	            
	              # @program.adjustments.destroy_all
	              @block_hash = {}
	              key = ''
	              (1..50).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row -1
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    elsif (c_i == 1)
	                      @block_hash[key][21] = value
	                    elsif (c_i == 2)
	                      @block_hash[key][30] = value
	                    elsif (c_i == 3)
	                      @block_hash[key][45] = value
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	            end
	            if  @block_hash.values.first.values.first == "21 Day"
              	@block_hash.shift
            	end
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end
  def import_jummbo6600_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 6600")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (10..35).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)

            	# term
            	@term = nil
              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
                @term = 30
              elsif @title.include?("20 Year")
                @term = 20
              elsif @title.include?("15 Year")
                @term = 15
              end
           
               	# interest type
              if @title.include?("Fixed")
                @rate_type = 0
              elsif @title.include?("ARM")
                @rate_type = 2
              else
              	@rate_type = nil
              end

              # streamline
              if @title.include?("FHA") 
                @streamline = true
                @fha = true
                @full_doc = true
              elsif @title.include?("VA")
              	@streamline = true
              	@va = true
              	@full_doc = true
              elsif @title.include?("USDA")
              	@streamline = true
              	@usda = true
              	@full_doc = true
              else
              	@streamline = nil
              	@full_doc = nil
              	@fha = nil
              	@va = nil
              	@usda = nil
              end

              # High Balance
              if @title.include?("High Bal")
              	@jumbo_high_balance = true
              else
              	@jumbo_high_balance = nil
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
                @rate_arm = @title.scan(/\d+/)[0].to_i
              else
              	@rate_arm = nil
              end
              
              @program = Program.find_or_create_by(program_name: @title)
              @programs_ids << @program.id
              @program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    elsif (c_i == 1)
                      @block_hash[key][21] = value
                    elsif (c_i == 2)
                      @block_hash[key][30] = value
                    elsif (c_i == 3)
                      @block_hash[key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end
  def import_jummbo7600_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 7600")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (10..36).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)

            	# term
            	@term = nil
              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
                @term = 30
              elsif @title.include?("20 Year")
                @term = 20
              elsif @title.include?("15 Year")
                @term = 15
              end
           
               	# interest type
              if @title.include?("Fixed")
                @rate_type = 0
              elsif @title.include?("ARM")
                @rate_type = 2
              else
              	@rate_type = nil
              end

              # streamline
              if @title.include?("FHA") 
                @streamline = true
                @fha = true
                @full_doc = true
              elsif @title.include?("VA")
              	@streamline = true
              	@va = true
              	@full_doc = true
              elsif @title.include?("USDA")
              	@streamline = true
              	@usda = true
              	@full_doc = true
              else
              	@streamline = nil
              	@full_doc = nil
              	@fha = nil
              	@va = nil
              	@usda = nil
              end

              # High Balance
              if @title.include?("High Bal")
              	@jumbo_high_balance = true
              else
              	@jumbo_high_balance = nil
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
                @rate_arm = @title.scan(/\d+/)[0].to_i
              else
              	@rate_arm = nil
              end
              
              @program = Program.find_or_create_by(program_name: @title)
              @programs_ids << @program.id
              @program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    elsif (c_i == 1)
                      @block_hash[key][21] = value
                    elsif (c_i == 2)
                      @block_hash[key][30] = value
                    elsif (c_i == 3)
                      @block_hash[key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end
  def import_jummbo6400_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 6400")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (10..41).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
              if @title.present? && cc < 9
	            	# term
	            	@term = nil
	              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
	                @term = 30
	              elsif @title.include?("20 Year")
	                @term = 20
	              elsif @title.include?("15 Year")
	                @term = 15
	              end
	           
	               	# interest type
	              if @title.include?("Fixed")
	                @rate_type = 0
	              elsif @title.include?("ARM")
	                @rate_type = 2
	              else
	              	@rate_type = nil
	              end

	              # streamline
	              if @title.include?("FHA") 
	                @streamline = true
	                @fha = true
	                @full_doc = true
	              elsif @title.include?("VA")
	              	@streamline = true
	              	@va = true
	              	@full_doc = true
	              elsif @title.include?("USDA")
	              	@streamline = true
	              	@usda = true
	              	@full_doc = true
	              else
	              	@streamline = nil
	              	@full_doc = nil
	              	@fha = nil
	              	@va = nil
	              	@usda = nil
	              end

	              # High Balance
	              if @title.include?("High Bal")
	              	@jumbo_high_balance = true
	              else
	              	@jumbo_high_balance = nil
	              end

	              # interest sub type
	              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
	                @rate_arm = @title.scan(/\d+/)[0].to_i
	              else
	              	@rate_arm = nil
	              end
	            end
              if @title.present? && cc < 9
	              @program = Program.find_or_create_by(program_name: @title)
	              @programs_ids << @program.id
	              @program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
	              # @program.adjustments.destroy_all
	              
	              @block_hash = {}
	              key = ''
	              (1..50).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row -1
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    elsif (c_i == 1)
	                      @block_hash[key][21] = value
	                    elsif (c_i == 2)
	                      @block_hash[key][30] = value
	                    elsif (c_i == 3)
	                      @block_hash[key][45] = value
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	            end
	            if @block_hash.keys.first == "Rate"
              	@block_hash.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        (44..58).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
          		if @title.present? && @title == "10/1 ARM - 6410"
	            	# term
	            	@term = nil
	              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
	                @term = 30
	              elsif @title.include?("20 Year")
	                @term = 20
	              elsif @title.include?("15 Year")
	                @term = 15
	              end
	           
	               	# interest type
	              if @title.include?("Fixed")
	                @rate_type = 0
	              elsif @title.include?("ARM")
	                @rate_type = 2
	              else
	              	@rate_type = nil
	              end

	              # streamline
	              if @title.include?("FHA") 
	                @streamline = true
	                @fha = true
	                @full_doc = true
	              elsif @title.include?("VA")
	              	@streamline = true
	              	@va = true
	              	@full_doc = true
	              elsif @title.include?("USDA")
	              	@streamline = true
	              	@usda = true
	              	@full_doc = true
	              else
	              	@streamline = nil
	              	@full_doc = nil
	              	@fha = nil
	              	@va = nil
	              	@usda = nil
	              end

	              # High Balance
	              if @title.include?("High Bal")
	              	@jumbo_high_balance = true
	              else
	              	@jumbo_high_balance = nil
	              end

	              # interest sub type
	              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
	                @rate_arm = @title.scan(/\d+/)[0].to_i
	              else
	              	@rate_arm = nil
	              end
              end
              if cc < 5 && @title == "10/1 ARM - 6410"
	              @program = Program.find_or_create_by(program_name: @title)
	              @programs_ids << @program.id
	             	@program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
	            
	              # @program.adjustments.destroy_all
	              @block_hash = {}
	              key = ''
	              (1..50).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row -1
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    elsif (c_i == 1)
	                      @block_hash[key][21] = value
	                    elsif (c_i == 2)
	                      @block_hash[key][30] = value
	                    elsif (c_i == 3)
	                      @block_hash[key][45] = value
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	            end
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end
  def import_jummbo6800_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 6800")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (10..37).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)

            	# term
            	@term = nil
              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
                @term = 30
              elsif @title.include?("20 Year")
                @term = 20
              elsif @title.include?("15 Year")
                @term = 15
              end
           
               	# interest type
              if @title.include?("Fixed")
                @rate_type = 0
              elsif @title.include?("ARM")
                @rate_type = 2
              else
              	@rate_type = nil
              end

              # streamline
              if @title.include?("FHA") 
                @streamline = true
                @fha = true
                @full_doc = true
              elsif @title.include?("VA")
              	@streamline = true
              	@va = true
              	@full_doc = true
              elsif @title.include?("USDA")
              	@streamline = true
              	@usda = true
              	@full_doc = true
              else
              	@streamline = nil
              	@full_doc = nil
              	@fha = nil
              	@va = nil
              	@usda = nil
              end

              # High Balance
              if @title.include?("High Bal")
              	@jumbo_high_balance = true
              else
              	@jumbo_high_balance = nil
              end

              # interest sub type
              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM") || @title.include?("5/1 LIBOR ARM") || @title.include?("7/1 LIBOR ARM") || @title.include?("10/1 LIBOR ARM")
                @rate_arm = @title.scan(/\d+/)[0].to_i
              else
              	@rate_arm = nil
              end
              
              @program = Program.find_or_create_by(program_name: @title)
              @programs_ids << @program.id
              @program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    elsif (c_i == 1)
                      @block_hash[key][21] = value
                    elsif (c_i == 2)
                      @block_hash[key][30] = value
                    elsif (c_i == 3)
                      @block_hash[key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @block_hash.shift
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end
  def import_jumbo6900_7900_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 6900 & 7900")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        (10..23).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
              	if @title.present?
		            	# term
		            	@term = nil
		              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
		                @term = 30
		              elsif @title.include?("20 Year")
		                @term = 20
		              elsif @title.include?("15 Year")
		                @term = 15
		              end
		           
		               	# interest type
		              if @title.include?("Fixed")
		                @rate_type = 0
		              elsif @title.include?("ARM")
		                @rate_type = 2
		              else
		              	@rate_type = nil
		              end

		              # streamline
		              if @title.include?("FHA") 
		                @streamline = true
		                @fha = true
		                @full_doc = true
		              elsif @title.include?("VA")
		              	@streamline = true
		              	@va = true
		              	@full_doc = true
		              elsif @title.include?("USDA")
		              	@streamline = true
		              	@usda = true
		              	@full_doc = true
		              else
		              	@streamline = nil
		              	@full_doc = nil
		              	@fha = nil
		              	@va = nil
		              	@usda = nil
		              end

		              # High Balance
		              if @title.include?("High Bal")
		              	@jumbo_high_balance = true
		              else
		              	@jumbo_high_balance = nil
		              end

		              # interest sub type
		              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM") || @title.include?("5/1 LIBOR ARM") || @title.include?("7/1 LIBOR ARM") || @title.include?("10/1 LIBOR ARM")
		                @rate_arm = @title.scan(/\d+/)[0].to_i
		              else
		              	@rate_arm = nil
		              end
              	end
	              @program = Program.find_or_create_by(program_name: @title)
	              @programs_ids << @program.id
	             	@program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
	            if @title.present?
	              # @program.adjustments.destroy_all
	              @block_hash = {}
	              key = ''
	              (1..50).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row -1
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    elsif (c_i == 1)
	                      @block_hash[key][21] = value
	                    elsif (c_i == 2)
	                      @block_hash[key][30] = value
	                    elsif (c_i == 3)
	                      @block_hash[key][45] = value
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	            end
	            if @block_hash.values.first.values.first == "21 Day"
              	@block_hash.shift
            	end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        (51..64).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
          	rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
              	if @title.present?
		            	# term
		            	@term = nil
		              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
		                @term = 30
		              elsif @title.include?("20 Year")
		                @term = 20
		              elsif @title.include?("15 Year")
		                @term = 15
		              end
		           
		               	# interest type
		              if @title.include?("Fixed")
		                @rate_type = 0
		              elsif @title.include?("ARM")
		                @rate_type = 2
		              else
		              	@rate_type = nil
		              end

		              # streamline
		              if @title.include?("FHA") 
		                @streamline = true
		                @fha = true
		                @full_doc = true
		              elsif @title.include?("VA")
		              	@streamline = true
		              	@va = true
		              	@full_doc = true
		              elsif @title.include?("USDA")
		              	@streamline = true
		              	@usda = true
		              	@full_doc = true
		              else
		              	@streamline = nil
		              	@full_doc = nil
		              	@fha = nil
		              	@va = nil
		              	@usda = nil
		              end

		              # High Balance
		              if @title.include?("High Bal")
		              	@jumbo_high_balance = true
		              else
		              	@jumbo_high_balance = nil
		              end

		              # interest sub type
		              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM") || @title.include?("5/1 LIBOR ARM") || @title.include?("7/1 LIBOR ARM") || @title.include?("10/1 LIBOR ARM")
		                @rate_arm = @title.scan(/\d+/)[0].to_i
		              else
		              	@rate_arm = nil
		              end
              	end
	              @program = Program.find_or_create_by(program_name: @title)
	              @programs_ids << @program.id
	             	@program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
	            if @title.present?
	              # @program.adjustments.destroy_all
	              @block_hash = {}
	              key = ''
	              (1..50).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row -1
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    elsif (c_i == 1)
	                      @block_hash[key][21] = value
	                    elsif (c_i == 2)
	                      @block_hash[key][30] = value
	                    elsif (c_i == 3)
	                      @block_hash[key][45] = value
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	            end
	            if @block_hash.values.first.values.first == "21 Day"
              	@block_hash.shift
            	end
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    # redirect_to programs_import_file_path(@bank)
  	redirect_to root_path
  end

  private
  def get_sheet
  	@sheet = Sheet.find(params[:id])
  end

  # def programs
  #   @programs = @sheet.programs.where(sheet_name: params[:sheet])
  # end

  private
  def get_value value1
    if value1.present?
      if value1.include?("FICO")
        value1 = value1.split("FICO ").last
      elsif value1.include?("<")
        value1 = "0"+value1
      elsif value1.include?("<=")
        value1 = "0"+value1
      else
        value1
      end
    end
  end
  # def get_sheet
  # 	debugger
  # 	@sheet = Sheet.find(params[:id])
  # end
  # def get_bank
  # 	debugger
  #   @bank = Bank.find(params[:id])
  # end
end
