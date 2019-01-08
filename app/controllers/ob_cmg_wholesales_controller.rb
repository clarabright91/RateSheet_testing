class ObCmgWholesalesController < ApplicationController
  def index
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
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
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
              end

              # High Balance
              if @title.include?("High Bal")
              	@jumbo_high_balance = true
              end

              @program = Program.find_or_create_by(title: @title)
              @programs_ids << @program.id
              @program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc)
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
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              else
              	@interest_type = nil
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

              @program = Program.find_or_create_by(title: @title)
              @programs_ids << @program.id
              @program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance)
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
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              else
              	@interest_type = nil
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
              
              @program = Program.find_or_create_by(title: @title)
              @programs_ids << @program.id
              @program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, interest_subtype: @interest_subtype)
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
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              else
              	@interest_type = nil
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
              
              @program = Program.find_or_create_by(title: @title)
              @programs_ids << @program.id
              @program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, interest_subtype: @interest_subtype)
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
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              else
              	@interest_type = nil
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
                @interest_subtype = @title.scan(/\d+/)[0].to_i
              else
              	@interest_subtype = nil
              end
              
              @program = Program.find_or_create_by(title: @title)
              @programs_ids << @program.id
              @program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, interest_subtype: @interest_subtype)
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
	                @interest_type = 0
	              elsif @title.include?("ARM")
	                @interest_type = 2
	              else
	              	@interest_type = nil
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
	                @interest_subtype = @title.scan(/\d+/)[0].to_i
	              else
	              	@interest_subtype = nil
	              end
              end
              
              if cc < 5
	              @program = Program.find_or_create_by(title: @title)
	              @programs_ids << @program.id
	             	@program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, interest_subtype: @interest_subtype)
	            
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
	                @interest_type = 0
	              elsif @title.include?("ARM")
	                @interest_type = 2
	              else
	              	@interest_type = nil
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
	                @interest_subtype = @title.scan(/\d+/)[0].to_i
	              else
	              	@interest_subtype = nil
	              end
              end
              
              if cc < 5
	              @program = Program.find_or_create_by(title: @title)
	              @programs_ids << @program.id
	             	@program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, interest_subtype: @interest_subtype)
	            
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
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              else
              	@interest_type = nil
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
                @interest_subtype = @title.scan(/\d+/)[0].to_i
              else
              	@interest_subtype = nil
              end
              
              @program = Program.find_or_create_by(title: @title)
              @programs_ids << @program.id
              @program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, interest_subtype: @interest_subtype)
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
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              else
              	@interest_type = nil
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
                @interest_subtype = @title.scan(/\d+/)[0].to_i
              else
              	@interest_subtype = nil
              end
              
              @program = Program.find_or_create_by(title: @title)
              @programs_ids << @program.id
              @program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, interest_subtype: @interest_subtype)
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
        (10..58).each do |r|
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
	                @interest_type = 0
	              elsif @title.include?("ARM")
	                @interest_type = 2
	              else
	              	@interest_type = nil
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
	                @interest_subtype = @title.scan(/\d+/)[0].to_i
	              else
	              	@interest_subtype = nil
	              end
	            end
              if @title.present? && cc < 9 && r != 43
	              @program = Program.find_or_create_by(title: @title)
	              @programs_ids << @program.id
	              @program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, interest_subtype: @interest_subtype)
	              # @program.adjustments.destroy_all
	              debugger
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
                @interest_type = 0
              elsif @title.include?("ARM")
                @interest_type = 2
              else
              	@interest_type = nil
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
                @interest_subtype = @title.scan(/\d+/)[0].to_i
              else
              	@interest_subtype = nil
              end
              
              @program = Program.find_or_create_by(title: @title)
              @programs_ids << @program.id
              @program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, interest_subtype: @interest_subtype)
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
		                @interest_type = 0
		              elsif @title.include?("ARM")
		                @interest_type = 2
		              else
		              	@interest_type = nil
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
		                @interest_subtype = @title.scan(/\d+/)[0].to_i
		              else
		              	@interest_subtype = nil
		              end
              	end
	              @program = Program.find_or_create_by(title: @title)
	              @programs_ids << @program.id
	             	@program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, interest_subtype: @interest_subtype)
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
		                @interest_type = 0
		              elsif @title.include?("ARM")
		                @interest_type = 2
		              else
		              	@interest_type = nil
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
		                @interest_subtype = @title.scan(/\d+/)[0].to_i
		              else
		              	@interest_subtype = nil
		              end
              	end
	              @program = Program.find_or_create_by(title: @title)
	              @programs_ids << @program.id
	             	@program.update(term: @term,interest_type: 0,loan_type: 0,streamline: @streamline,fha: @fha, va: @va, usda: @usda, full_doc: @full_doc, jumbo_high_balance: @jumbo_high_balance, interest_subtype: @interest_subtype)
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
end
