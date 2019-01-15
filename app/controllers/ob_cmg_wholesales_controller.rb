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
              	@full_doc = false
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
        @adjustment_hash = {}
        primary_key = ''
        primary_key1 = ''
        secondary_key = ''
        fnma_key = ''
        sub_data = ''
        sub_key = ''
        cltv_key = ''
        cap_key = ''
        m_key = ''
        key = ''
        loan_key = ''
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
        # Adjustment
        (55..88).each do |r|
        	row = sheet_data.row(r)
        	@fnma_data = sheet_data.row(57)
        	@sub_data = sheet_data.row(73)
        	@cap_data = sheet_data.row(80)
        	if row.compact.count >= 1
        		(0..16).each do |cc|
        			value = sheet_data.cell(r,cc)
        			if value.present?

        				if value == "FNMA DU REFI PLUS ADJUSTMENTS"
        					primary_key = "FNMA/DU"
        					@adjustment_hash[primary_key] = {}
        				elsif value == "SUBORDINATE FINANCING"
        					primary_key = "FinancingType/LTV/CLTV/FICO"
        					sub_key = "Subordinate Financing"
        					@adjustment_hash[primary_key] = {}
        					@adjustment_hash[primary_key][sub_key] = {}
        				elsif value == "DU REFI PLUS ADJUSTMENT CAP (MAX ADJ) *"
									primary_key = value
									@adjustment_hash[primary_key] = {}
        				end
        				if r >= 58 && r <= 70 && cc == 1
        					secondary_key = get_value value
        					@adjustment_hash[primary_key][secondary_key] = {}
        				end
        				if r >= 58 && r <= 70 && cc >= 8 && cc <= 16
        					fnma_key = get_value @fnma_data[cc-1]
        					@adjustment_hash[primary_key][secondary_key][fnma_key] = {}
        					@adjustment_hash[primary_key][secondary_key][fnma_key] = value
        				end

        				# subordinate adjustment
        				if r >= 74 && r <= 78 && cc == 1
        					secondary_key = get_value value
        					@adjustment_hash[primary_key][sub_key][secondary_key] = {}
        				end
        				if r >= 74 && r <= 78 && cc == 3
        					cltv_key = get_value value
        					@adjustment_hash[primary_key][sub_key][secondary_key][cltv_key] = {}
        				end
        				if r >= 74 && r <= 78 && cc >= 5 && cc <= 7
        					sub_data = get_value @sub_data[cc-1]
        					@adjustment_hash[primary_key][sub_key][secondary_key][cltv_key][sub_data] = {}
        					@adjustment_hash[primary_key][sub_key][secondary_key][cltv_key][sub_data] = value
        				end
        				# Adjustment Cap
        				if r >= 81 && r <= 83 && cc == 1
        					secondary_key = value
        					@adjustment_hash[primary_key][secondary_key] = {}
        				end
        				if r >= 81 && r <= 83 && cc == 4
        					cltv_key = get_value value
        					@adjustment_hash[primary_key][secondary_key][cltv_key] = {}
        				end
        				if r >= 81 && r <= 83 && cc >= 5 && cc <= 7
        					cap_key = get_value @cap_data[cc-1]
        					@adjustment_hash[primary_key][secondary_key][cltv_key][cap_key] = {}
        					@adjustment_hash[primary_key][secondary_key][cltv_key][cap_key] = value
        				end
        			end
        		end
        		(10..16).each do |cc|
        			value = sheet_data.cell(r,cc)
        			if value == "MISCELLANEOUS"
        				primary_key1 = "Miscellaneous"
        				@adjustment_hash[primary_key1] = {}
        			end
        			if value == "LOAN AMOUNT "
        				primary_key1 = "RateType/LoanAmount/CLTV"
        				@adjustment_hash[primary_key1] = {}
        			end
        			if value == "STATE ADJUSTMENTS"
        				primary_key1 = "State"
        				@adjustment_hash[primary_key1] = {}
        			end
        			if value.present?
        				# MISCELLANEOUS
        				if r >= 73 && r <= 74 && cc == 10
        					m_key = value
        					@adjustment_hash[primary_key1][m_key] = {}
        				end
        				if r >= 73 && r <= 74 && cc == 16
        					@adjustment_hash[primary_key1][m_key] = value
        				end
        				# LOAN AMOUNT ADJUSTMENT
        				if r >= 76 && r <= 80 && cc == 10
        					# m_key = value
        					m_key =  value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
        					@adjustment_hash[primary_key1][m_key] = {}
        				end
        				if r >= 76 && r <= 80 && cc == 16
        					@adjustment_hash[primary_key1][m_key] = value
        				end
        				# STATE ADJUSTMENTS
        				if r >= 83 && r <= 88 && cc == 11
        					adj_key = value.split(', ')
                  adj_key.each do |f_key|
                    key = f_key
                    ccc = cc + 5
                    c_val = sheet_data.cell(r,ccc)
                    @adjustment_hash[primary_key1][key] = c_val
                  end
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
  def import_oa_sheet
    @programs_ids = []
    file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "OA")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        primary_key = ''
        primary_key1 = ''
        secondary_key = ''
        fnma_key = ''
        sub_data = ''
        sub_key = ''
        cltv_key = ''
        cap_key = ''
        m_key = ''
        key = ''
        loan_key = ''
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
        # Adjustment
        (54..88).each do |r|
        	row = sheet_data.row(r)
        	@fnma_data = sheet_data.row(56)
        	@sub_data = sheet_data.row(73)
        	@cap_data = sheet_data.row(82)
        	if row.compact.count >= 1
        		(0..16).each do |cc|
        			value = sheet_data.cell(r,cc)
        			if value.present?

        				if value == "FHLMC LP OPEN ACCESS ADJUSTMENTS"
        					primary_key = "FHLMC/LP"
        					@adjustment_hash[primary_key] = {}
        				elsif value == "SUBORDINATE FINANCING"
        					primary_key = "FinancingType/LTV/CLTV/FICO"
        					sub_key = "Subordinate Financing"
        					@adjustment_hash[primary_key] = {}
        					@adjustment_hash[primary_key][sub_key] = {}
        				elsif value == "OPEN ACCESS ADJUSTMENT CAP (MAX ADJ) *"
									primary_key = value
									@adjustment_hash[primary_key] = {}
        				end
        				if r >= 57 && r <= 70 && cc == 1
        					secondary_key = get_value value
        					@adjustment_hash[primary_key][secondary_key] = {}
        				end
        				if r >= 57 && r <= 70 && cc >= 8 && cc <= 16
        					fnma_key = get_value @fnma_data[cc-1]
        					@adjustment_hash[primary_key][secondary_key][fnma_key] = {}
        					@adjustment_hash[primary_key][secondary_key][fnma_key] = value
        				end

        				# subordinate adjustment
        				if r >= 74 && r <= 80 && cc == 1
        					secondary_key = get_value value
        					@adjustment_hash[primary_key][sub_key][secondary_key] = {}
        				end
        				if r >= 74 && r <= 80 && cc == 3
        					cltv_key = get_value value
        					@adjustment_hash[primary_key][sub_key][secondary_key][cltv_key] = {}
        				end
        				if r >= 74 && r <= 80 && cc >= 5 && cc <= 7
        					sub_data = get_value @sub_data[cc-1]
        					@adjustment_hash[primary_key][sub_key][secondary_key][cltv_key][sub_data] = {}
        					@adjustment_hash[primary_key][sub_key][secondary_key][cltv_key][sub_data] = value
        				end
        				# Adjustment Cap
        				if r >= 83 && r <= 85 && cc == 1
        					secondary_key = value
        					@adjustment_hash[primary_key][secondary_key] = {}
        				end
        				if r >= 83 && r <= 85 && cc == 4
        					cltv_key = get_value value
        					@adjustment_hash[primary_key][secondary_key][cltv_key] = {}
        				end
        				if r >= 83 && r <= 85 && cc >= 5 && cc <= 7
        					cap_key = get_value @cap_data[cc-1]
        					@adjustment_hash[primary_key][secondary_key][cltv_key][cap_key] = {}
        					@adjustment_hash[primary_key][secondary_key][cltv_key][cap_key] = value
        				end
        			end
        		end
        		(10..16).each do |cc|
        			value = sheet_data.cell(r,cc)
        			if value.present?
        				if value == "MISCELLANEOUS"
	        				primary_key1 = "Miscellaneous"
	        				@adjustment_hash[primary_key1] = {}
	        			end
	        			if value == "LOAN AMOUNT "
	        				primary_key1 = "RateType/LoanAmount/CLTV"
	        				@adjustment_hash[primary_key1] = {}
	        			end
	        			if value == "STATE ADJUSTMENTS"
	        				primary_key1 = "State"
	        				@adjustment_hash[primary_key1] = {}
	        			end

        				# MISCELLANEOUS
        				if r >= 73 && r <= 74 && cc == 10
        					m_key = value
        					@adjustment_hash[primary_key1][m_key] = {}
        				end
        				if r >= 73 && r <= 74 && cc == 16
        					@adjustment_hash[primary_key1][m_key] = value
        				end
        				# LOAN AMOUNT ADJUSTMENT
        				if r >= 76 && r <= 80 && cc == 10
        					m_key =  value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
        					@adjustment_hash[primary_key1][m_key] = {}
        				end
        				if r >= 76 && r <= 80 && cc == 16
        					@adjustment_hash[primary_key1][m_key] = value
        				end
        				# STATE ADJUSTMENTS
        				if r >= 83 && r <= 88 && cc == 11
        					adj_key = value.split(', ')
                  adj_key.each do |f_key|
                    key = f_key
                    ccc = cc + 5
                    c_val = sheet_data.cell(r,ccc)
                    @adjustment_hash[primary_key1][key] = c_val
                  end
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
        @adjustment_hash = {}
        @purchase_adjustment = {}
        @rate_adjustment = {}
        @other_adjustment = {}
        @jumbo_purchase_adjustment = {}
        @jumbo_rate_adjustment = {}
        @jumbo_other_adjustment = {}
        @cltv_data = []
        @ltv_data = []
        primary_key = ''
        secondary_key = ''
        cltv_key = ''
        m_key = ''
        max_key = ''
        ltv_key = ''
        key = ''
        adj_key = ''
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
        (12..46).each do |r|
        	row = sheet_data.row(r)
        	@cltv_data = sheet_data.row(13)
        	if row.compact.count >= 1
        		(6..16).each do |cc|
        			value = sheet_data.cell(r,cc)
        			if value.present?
        				if value == "Purchase Transaction"
        					primary_key = "LoanType/LTV/FICO"
        					@purchase_adjustment[primary_key] = {}
        				elsif value == "Rate/Term Transaction"
        					primary_key = "RateType/Term/LTV/FICO"
        					@rate_adjustment[primary_key] = {}
        				elsif value == "Cash Out Transaction"
        					primary_key = "LoanType/RefinanceOption/LTV"
        					@adjustment_hash[primary_key] = {}
        				end
        				# Purchase Transaction Adjustment
        				if r >= 14 && r <= 19 && cc == 6
        					secondary_key = get_value value
        					@purchase_adjustment[primary_key][secondary_key] = {}
        				end
        				if r >= 14 && r <= 19 && cc >= 10 && cc <= 16
        					cltv_key = get_value @cltv_data[cc-1]
        					@purchase_adjustment[primary_key][secondary_key][cltv_key] = {}
        					@purchase_adjustment[primary_key][secondary_key][cltv_key] = value
        				end

        				# Rate/Term Transaction Adjustment
        				if r >= 22 && r <= 27 && cc == 6
        					secondary_key = get_value value
        					@rate_adjustment[primary_key][secondary_key] = {}
        				end
        				if r >= 22 && r <= 27 && cc >= 10 && cc <= 16
        					cltv_key = get_value @cltv_data[cc-1]
        					@rate_adjustment[primary_key][secondary_key][cltv_key] = {}
        					@rate_adjustment[primary_key][secondary_key][cltv_key] = value
        				end

        				# Cash Out Transaction Adjustment
        				if r >= 30 && r <= 46 && cc == 6
        					if value.include?("Loan Amount")
        						secondary_key = value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
        					else
        						secondary_key = get_value value
        					end
        					@adjustment_hash[primary_key][secondary_key] = {}
        				end
        				if r >= 30 && r <= 48 && cc >= 10 && cc <= 16
        					cltv_key = get_value @cltv_data[cc-1]
        					@adjustment_hash[primary_key][secondary_key][cltv_key] = {}
        					@adjustment_hash[primary_key][secondary_key][cltv_key] = value
        				end
        			end
        		end
        		(1..4).each do |cc|
        			value = sheet_data.cell(r,cc)
        			if value.present?
        				if value == "MAX PRICE AFTER ADJUSTMENTS"
        					max_key = "RateType/LA/"
        					@other_adjustment[max_key] = {}
        				end
        				# MISCELLANEOUS
	        			if r == 25 && cc == 1
	        				m_key = "Miscellaneous/NY"
	        				@other_adjustment[m_key] = {}
	        			end
	        			if r == 25 && cc == 4
	        				@other_adjustment[m_key] = value
	        			end
	        			# MAX PRICE AFTER ADJUSTMENTS
	        			if r >= 29 && r <= 30 && cc == 1
	        				key = value.include?("LA <") ? "0" + value.split("LA").last : value.split("LA").last
        					@other_adjustment[max_key][key] = {}
	        			end
	        			if r >= 29 && r <= 30 && cc == 4
	        				@other_adjustment[max_key][key] = value
	        			end
	        		end
        		end
        	end
        end
        Adjustment.create(data: @purchase_adjustment, sheet_name: sheet)
        Adjustment.create(data: @rate_adjustment, sheet_name: sheet)
        Adjustment.create(data: @adjustment_hash, sheet_name: sheet)
        Adjustment.create(data: @other_adjustment, sheet_name: sheet)
        (56..77).each do |r|
        	row = sheet_data.row(r)
        	@ltv_data = sheet_data.row(59)
        	if row.compact.count >= 1
        		(10..16).each do |cc|
        			value = sheet_data.cell(r,cc)
        			if value.present?
        				if value == "Purchase Transaction"
        					primary_key = "Jumbo/LoanType/LTV/FICO"
        					@jumbo_purchase_adjustment[primary_key] = {}
        				elsif value == "Rate/Term Transaction"
        					primary_key = "Jumbo/RateType/Term/LTV/FICO"
        					@jumbo_rate_adjustment[primary_key] = {}
        				elsif value == "MISCELLANEOUS"
        					primary_key = "Jumbo/NY/LTV/FICO"
        					@jumbo_other_adjustment[primary_key] = {}
        				end
        				# Purchase Transaction Adjustment
        				if r >= 60 && r <= 63 && cc == 10
        					secondary_key = get_value value
        					@jumbo_purchase_adjustment[primary_key][secondary_key] = {}
        				end
        				if r >= 60 && r <= 63 && cc >= 15 && cc <= 16
        					ltv_key = get_value @ltv_data[cc-1]
        					@jumbo_purchase_adjustment[primary_key][secondary_key][ltv_key] = {}
        					@jumbo_purchase_adjustment[primary_key][secondary_key][ltv_key] = value
        				end

        				# Rate/Term Transaction Adjustment
        				if r >= 66 && r <= 71 && cc == 10
        					if value.include?("Loan Amount")
        						secondary_key = value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
        					else
        						secondary_key = get_value value
        					end
        					@jumbo_rate_adjustment[primary_key][secondary_key] = {}
        				end
        				if r >= 66 && r <= 71 && cc >= 15 && cc <= 16
        					ltv_key = get_value @ltv_data[cc-1]
        					@jumbo_rate_adjustment[primary_key][secondary_key][ltv_key] = {}
        					@jumbo_rate_adjustment[primary_key][secondary_key][ltv_key] = value
        				end

        				if r == 72 && cc == 10
        					key = "FL"
        					adj_key = "NV"
        					@jumbo_rate_adjustment[primary_key][key] = {}
        					@jumbo_rate_adjustment[primary_key][adj_key] = {}
        				end
        				if r == 72 && cc >= 15 && cc <= 16
        					@jumbo_rate_adjustment[primary_key][key][@ltv_data[cc-1]] = {}
        					@jumbo_rate_adjustment[primary_key][adj_key][@ltv_data[cc-1]] = {}
        					@jumbo_rate_adjustment[primary_key][key][@ltv_data[cc-1]] = value
        					@jumbo_rate_adjustment[primary_key][adj_key][@ltv_data[cc-1]] = value
        				end
        				if r >= 73 && r <= 74 && cc == 10
        					secondary_key = get_value value
        					@jumbo_rate_adjustment[primary_key][secondary_key] = {}
        				end
        				if r >= 73 && r <= 74 && cc >= 15 && cc <= 16
        					@jumbo_rate_adjustment[primary_key][secondary_key][@ltv_data[cc-1]] = {}
        					@jumbo_rate_adjustment[primary_key][secondary_key][@ltv_data[cc-1]] =  value
        				end
        				# MISCELLANEOUS
        				if r == 77 && cc == 10
        					m_key = value
        					@jumbo_other_adjustment[primary_key][m_key] = {}
        				end
        				if r == 77 && cc == 16
        					@jumbo_other_adjustment[primary_key][m_key] = value
        				end
        			end
        		end
        		(0..3).each do |cc|
        			value = sheet_data.cell(r,cc)
        			if value.present?
        				if value == "MAX PRICE AFTER ADJUSTMENTS"
        					max_key = "RateType/LA/"
        					@jumbo_other_adjustment[max_key] = {}
        				end
        				if r >= 71 && r <= 72 && cc == 1
        					key = value.include?("LA <") ? "0" + value.split("LA").last : value.split("LA").last
        					@jumbo_other_adjustment[max_key][key] = {}
        				end
        				if r >= 71 && r <= 72 && cc == 3
        					@jumbo_other_adjustment[max_key][key] = value
        				end
        			end
        		end
        	end
        end
        Adjustment.create(data: @jumbo_purchase_adjustment, sheet_name: sheet)
        Adjustment.create(data: @jumbo_rate_adjustment, sheet_name: sheet)
        Adjustment.create(data: @jumbo_other_adjustment, sheet_name: sheet)
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
  # def import_jummbo6400_sheet
  #   @programs_ids = []
  #   file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
  #   xlsx = Roo::Spreadsheet.open(file)
  #   xlsx.sheets.each do |sheet|
  #     if (sheet == "JUMBO 6400")
  #       sheet_data = xlsx.sheet(sheet)
  #       @programs_ids = []
  #       (10..41).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 4))
  #         	rr = r + 1
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 4*max_column + 1

  #             @title = sheet_data.cell(r,cc)
  #             if @title.present? && cc < 9
	 #            	# term
	 #            	@term = nil
	 #              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
	 #                @term = 30
	 #              elsif @title.include?("20 Year")
	 #                @term = 20
	 #              elsif @title.include?("15 Year")
	 #                @term = 15
	 #              end
	           
	 #               	# interest type
	 #              if @title.include?("Fixed")
	 #                @rate_type = 0
	 #              elsif @title.include?("ARM")
	 #                @rate_type = 2
	 #              else
	 #              	@rate_type = nil
	 #              end

	 #              # streamline
	 #              if @title.include?("FHA") 
	 #                @streamline = true
	 #                @fha = true
	 #                @full_doc = true
	 #              elsif @title.include?("VA")
	 #              	@streamline = true
	 #              	@va = true
	 #              	@full_doc = true
	 #              elsif @title.include?("USDA")
	 #              	@streamline = true
	 #              	@usda = true
	 #              	@full_doc = true
	 #              else
	 #              	@streamline = nil
	 #              	@full_doc = nil
	 #              	@fha = nil
	 #              	@va = nil
	 #              	@usda = nil
	 #              end

	 #              # High Balance
	 #              if @title.include?("High Bal")
	 #              	@jumbo_high_balance = true
	 #              else
	 #              	@jumbo_high_balance = nil
	 #              end

	 #              # interest sub type
	 #              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
	 #                @rate_arm = @title.scan(/\d+/)[0].to_i
	 #              else
	 #              	@rate_arm = nil
	 #              end
	 #            end
  #             if @title.present? && cc < 9
	 #              @program = Program.find_or_create_by(program_name: @title)
	 #              @programs_ids << @program.id
	 #              @program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
	 #              # @program.adjustments.destroy_all
	              
	 #              @block_hash = {}
	 #              key = ''
	 #              (1..50).each do |max_row|
	 #                @data = []
	 #                (0..3).each_with_index do |index, c_i|
	 #                  rrr = rr + max_row -1
	 #                  ccc = cc + c_i
	 #                  value = sheet_data.cell(rrr,ccc)
	 #                  if value.present?
	 #                    if (c_i == 0)
	 #                      key = value
	 #                      @block_hash[key] = {}
	 #                    elsif (c_i == 1)
	 #                      @block_hash[key][21] = value
	 #                    elsif (c_i == 2)
	 #                      @block_hash[key][30] = value
	 #                    elsif (c_i == 3)
	 #                      @block_hash[key][45] = value
	 #                    end
	 #                    @data << value
	 #                  end
	 #                end
	 #                if @data.compact.reject { |c| c.blank? }.length == 0
	 #                  break # terminate the loop
	 #                end
	 #              end
	 #            end
	 #            if @block_hash.keys.first == "Rate"
  #             	@block_hash.shift
  #             end
  #             @program.update(base_rate: @block_hash)
  #           end
  #         end
  #       end
  #       (44..58).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 4))
  #         	rr = r + 1
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 4*max_column + 1

  #             @title = sheet_data.cell(r,cc)
  #         		if @title.present? && @title == "10/1 ARM - 6410"
	 #            	# term
	 #            	@term = nil
	 #              if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
	 #                @term = 30
	 #              elsif @title.include?("20 Year")
	 #                @term = 20
	 #              elsif @title.include?("15 Year")
	 #                @term = 15
	 #              end
	           
	 #               	# interest type
	 #              if @title.include?("Fixed")
	 #                @rate_type = 0
	 #              elsif @title.include?("ARM")
	 #                @rate_type = 2
	 #              else
	 #              	@rate_type = nil
	 #              end

	 #              # streamline
	 #              if @title.include?("FHA") 
	 #                @streamline = true
	 #                @fha = true
	 #                @full_doc = true
	 #              elsif @title.include?("VA")
	 #              	@streamline = true
	 #              	@va = true
	 #              	@full_doc = true
	 #              elsif @title.include?("USDA")
	 #              	@streamline = true
	 #              	@usda = true
	 #              	@full_doc = true
	 #              else
	 #              	@streamline = nil
	 #              	@full_doc = nil
	 #              	@fha = nil
	 #              	@va = nil
	 #              	@usda = nil
	 #              end

	 #              # High Balance
	 #              if @title.include?("High Bal")
	 #              	@jumbo_high_balance = true
	 #              else
	 #              	@jumbo_high_balance = nil
	 #              end

	 #              # interest sub type
	 #              if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
	 #                @rate_arm = @title.scan(/\d+/)[0].to_i
	 #              else
	 #              	@rate_arm = nil
	 #              end
  #             end
  #             if cc < 5 && @title == "10/1 ARM - 6410"
	 #              @program = Program.find_or_create_by(program_name: @title)
	 #              @programs_ids << @program.id
	 #              end
	 #            end
  #             if @title.present? && cc < 9
	 #              @program = Program.find_or_create_by(program_name: @title)
	 #              @programs_ids << @program.id
	 #              @program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
	 #              # @program.adjustments.destroy_all
	              
	 #              @block_hash = {}
	 #              key = ''
	 #              (1..50).each do |max_row|
	 #                @data = []
	 #                (0..3).each_with_index do |index, c_i|
	 #                  rrr = rr + max_row -1
	 #                  ccc = cc + c_i
	 #                  value = sheet_data.cell(rrr,ccc)
	 #                  if value.present?
	 #                    if (c_i == 0)
	 #                      key = value
	 #                      @block_hash[key] = {}
	 #                    elsif (c_i == 1)
	 #                      @block_hash[key][21] = value
	 #                    elsif (c_i == 2)
	 #                      @block_hash[key][30] = value
	 #                    elsif (c_i == 3)
	 #                      @block_hash[key][45] = value
	 #                    end
	 #                    @data << value
	 #                  end
	 #                end
	 #                if @data.compact.reject { |c| c.blank? }.length == 0
	 #                  break # terminate the loop
	 #                end
	 #              end
	 #            end
	 #            if @block_hash.keys.first == "Rate"
  #             	@block_hash.shift
  #             end
  #             @program.update(base_rate: @block_hash)
  #           end
  #         end
  #       end
  #       # (44..58).each do |r|
  #       #   row = sheet_data.row(r)
  #       #   if ((row.compact.count > 1) && (row.compact.count <= 4))
  #       #   	rr = r + 1
  #       #     max_column_section = row.compact.count - 1
  #       #     (0..max_column_section).each do |max_column|
  #       #       cc = 4*max_column + 1

  #       #       @title = sheet_data.cell(r,cc)
  #       #   		if @title.present? && @title == "10/1 ARM - 6410"
	 #       #      	# term
	 #       #      	@term = nil
	 #       #        if @title.include?("30 Year") || @title.include?("30Yr") || @title.include?("30 Yr")
	 #       #          @term = 30
	 #       #        elsif @title.include?("20 Year")
	 #       #          @term = 20
	 #       #        elsif @title.include?("15 Year")
	 #       #          @term = 15
	 #       #        end
	           
	 #       #         	# interest type
	 #       #        if @title.include?("Fixed")
	 #       #          @rate_type = 0
	 #       #        elsif @title.include?("ARM")
	 #       #          @rate_type = 2
	 #       #        else
	 #       #        	@rate_type = nil
	 #       #        end

	 #       #        # streamline
	 #       #        if @title.include?("FHA") 
	 #       #          @streamline = true
	 #       #          @fha = true
	 #       #          @full_doc = true
	 #       #        elsif @title.include?("VA")
	 #       #        	@streamline = true
	 #       #        	@va = true
	 #       #        	@full_doc = true
	 #       #        elsif @title.include?("USDA")
	 #       #        	@streamline = true
	 #       #        	@usda = true
	 #       #        	@full_doc = true
	 #       #        else
	 #       #        	@streamline = nil
	 #       #        	@full_doc = nil
	 #       #        	@fha = nil
	 #       #        	@va = nil
	 #       #        	@usda = nil
	 #       #        end

	 #       #        # High Balance
	 #       #        if @title.include?("High Bal")
	 #       #        	@jumbo_high_balance = true
	 #       #        else
	 #       #        	@jumbo_high_balance = nil
	 #       #        end

	 #       #        # interest sub type
	 #       #        if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
	 #       #          @rate_arm = @title.scan(/\d+/)[0].to_i
	 #       #        else
	 #       #        	@rate_arm = nil
	 #       #        end
  #       #       end
  #       #       if cc < 5 && @title == "10/1 ARM - 6410"
	 #       #        @program = Program.find_or_create_by(program_name: @title)
	 #       #        @programs_ids << @program.id
	 #       #       	@program.update(term: @term,rate_type: @rate_type,loan_type: "Purchase",streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, rate_arm: @rate_arm)
	            
	 #       #        # @program.adjustments.destroy_all
	 #       #        @block_hash = {}
	 #       #        key = ''
	 #       #        (1..50).each do |max_row|
	 #       #          @data = []
	 #       #          (0..3).each_with_index do |index, c_i|
	 #       #            rrr = rr + max_row -1
	 #       #            ccc = cc + c_i
	 #       #            value = sheet_data.cell(rrr,ccc)
	 #       #            if value.present?
	 #       #              if (c_i == 0)
	 #       #                key = value
	 #       #                @block_hash[key] = {}
	 #       #              elsif (c_i == 1)
	 #       #                @block_hash[key][21] = value
	 #       #              elsif (c_i == 2)
	 #       #                @block_hash[key][30] = value
	 #       #              elsif (c_i == 3)
	 #       #                @block_hash[key][45] = value
	 #       #              end
	 #       #              @data << value
	 #       #            end
	 #       #          end
	 #       #          if @data.compact.reject { |c| c.blank? }.length == 0
	 #       #            break # terminate the loop
	 #       #          end
	 #       #        end
	 #       #      end
  #       #       @block_hash.shift
  #       #       @program.update(base_rate: @block_hash)
  #       #     end
  #       #   end
  #       # end
  #     end
  #   end
  #   # redirect_to programs_import_file_path(@bank)
  # 	redirect_to root_path
  # end
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

  # private
  # def get_sheet
  # 	@sheet = Sheet.find(params[:id])
  # end
  def get_value value1
  	if value1.present?
  		if value1.include?("FICO <") 
  			value1 = "0"+value1.split("FICO").last
  		elsif value1.include?("<")
  			value1 = "0"+value1
     	elsif value1.include?("FICO")
       	value1 = value1.split("FICO ").last
      elsif value1 == "Investment Property"
      	value1 = "Property/Type"
     	else
       	value1
     	end
   	end
 	end
end
