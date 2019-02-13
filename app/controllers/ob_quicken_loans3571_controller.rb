class ObQuickenLoans3571Controller < ApplicationController
	before_action :get_sheet, only: [:programs, :ws_du_lp_pricing, :durp_lp_relief_pricing, :fha_usda_full_doc_pricing, :fha_streamline_pricing, :va_full_doc_pricing, :va_irrrl_pricing_govy_llpas, :na_jumbo_pricing_llpas, :du_lp_llpas, :durp_lp_relief_llpas]
  before_action :read_sheet, only: [:index,:ws_du_lp_pricing, :durp_lp_relief_pricing, :fha_usda_full_doc_pricing, :fha_streamline_pricing, :va_full_doc_pricing, :va_irrrl_pricing_govy_llpas, :na_jumbo_pricing_llpas, :du_lp_llpas, :durp_lp_relief_llpas]
  before_action :get_program, only: [:single_program, :program_property]

	def index
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "WS Rate Sheet Summary")
          # headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "Quicken Loans Mortgage Services"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def ws_du_lp_pricing
  	@xlsx.sheets.each do |sheet|
      if (sheet == "WS DU & LP Pricing")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (15..129).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5 
              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != "=FALSE()"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property              
	              @block_hash = {}
	              key = ''
	              (1..25).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    else
	                      @block_hash[key][15*c_i] = value if key.present?
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	              @program.update(base_rate: @block_hash)
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def durp_lp_relief_pricing
  	@xlsx.sheets.each do |sheet|
      if (sheet == "DURP & LP Relief Pricing")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (14..124).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5 
              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != "=FALSE()"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property              
	              @block_hash = {}
	              key = ''
	              (1..15).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    else
	                      @block_hash[key][15*c_i] = value if key.present?
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	              @program.update(base_rate: @block_hash)
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def fha_usda_full_doc_pricing
  	@xlsx.sheets.each do |sheet|
      if (sheet == "FHA & USDA Full Doc Pricing")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (13..93).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5 
              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != "=FALSE()"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property              
	              @block_hash = {}
	              key = ''
	              (1..20).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    else
	                      @block_hash[key][15*c_i] = value if key.present?
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	              @program.update(base_rate: @block_hash)
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def fha_streamline_pricing
  	@xlsx.sheets.each do |sheet|
      if (sheet == "FHA Streamline Pricing")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (13..93).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5 
              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != "=FALSE()"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property              
	              @block_hash = {}
	              key = ''
	              (1..20).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    else
	                      @block_hash[key][15*c_i] = value if key.present?
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	              @program.update(base_rate: @block_hash)
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def va_full_doc_pricing
  	@xlsx.sheets.each do |sheet|
      if (sheet == "VA Full Doc Pricing")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (13..88).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5 
              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != "=FALSE()"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property              
	              @block_hash = {}
	              key = ''
	              (1..20).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    else
	                      @block_hash[key][15*c_i] = value if key.present?
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	              @program.update(base_rate: @block_hash)
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def va_irrrl_pricing_govy_llpas
  	@xlsx.sheets.each do |sheet|
      if (sheet == "VA IRRRL Pricing & Govy LLPAs")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @government_hash = {}
        primary_key1 = ''
        secondary_key = ''

        # programs
        (13..66).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5 
              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != "=FALSE()"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property              
	              @block_hash = {}
	              key = ''
	              (1..20).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    else
	                      @block_hash[key][15*c_i] = value if key.present?
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	              @program.update(base_rate: @block_hash)
              end
            end
          end
        end

        (75..93).each do |r|
        	row = sheet_data.row(r)
        	(0..29).each do |cc|
        		value = sheet_data.cell(r,cc)
        		if value.present?
        			if value == "Government FICO Adjusters"
        				primary_key1 = "LoanType/FICO"
        				@government_hash[primary_key1] = {}
        			end
        			if r == 76 && cc == 5
        				primary_key = "LockDay"
        				@adjustment_hash[primary_key] = {}
        				if @adjustment_hash[primary_key] = {}
        					cc = cc +10
        					new_value = sheet_data.cell(r,cc)
        					@adjustment_hash[primary_key] = new_value
        				end
        			end
        			# Loan Ladder
        			if r >= 88 && r <= 91 && cc == 5
        				primary_key = get_value value
        				@adjustment_hash[primary_key] = {}
        				if @adjustment_hash[primary_key] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@adjustment_hash[primary_key] = new_value
        				end
        			end
        			# Government FICO Adjusters
        			if r >= 76 && r <= 79 && cc == 19
        				secondary_key = get_value value
        				@government_hash[primary_key1][secondary_key] = {}
        				if @government_hash[primary_key1][secondary_key] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@government_hash[primary_key1][secondary_key] = new_value
        				end
        			end
        			# Geography
        			if r == 82 && cc == 19
        				primary_key1 = "NJ"
        				@government_hash[primary_key1] = {}
        				if @government_hash[primary_key1] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@government_hash[primary_key1] = new_value
        				end
        			end
        			if r == 83 && cc == 19
        				primary_key1 = "NY"
        				@government_hash[primary_key1] = {}
        				if @government_hash[primary_key1] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@government_hash[primary_key1] = new_value
        				end
        			end
        			if r == 86 && cc == 19
        				primary_key1 = "VA/LoanPurpose/>=95%"
        				@government_hash[primary_key1] = {}
        				if @government_hash[primary_key1] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@government_hash[primary_key1] = new_value
        				end
        			end
        			if r == 87 && cc == 19
        				primary_key1 = "VA/RefinanceOption/>90%"
        				@government_hash[primary_key1] = {}
        				if @government_hash[primary_key1] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@government_hash[primary_key1] = new_value
        				end
        			end
        			if r == 88 && cc == 19
        				primary_key1 = "19 for 19 through 19th"
        				@government_hash[primary_key1] = {}
        				if @government_hash[primary_key1] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@government_hash[primary_key1] = new_value
        				end
        			end
        			if r == 88 && cc == 19
        				primary_key1 = "VA/>100%"
        				@government_hash[primary_key1] = {}
        				if @government_hash[primary_key1] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@government_hash[primary_key1] = new_value
        				end
        			end
        		end
        	end
        end
        adjustment = [@adjustment_hash,@government_hash]
        make_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def na_jumbo_pricing_llpas
  	@xlsx.sheets.each do |sheet|
      if (sheet == "NA Jumbo Pricing & LLPAs")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (7..50).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5 
              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != "=FALSE()"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property              
	              @block_hash = {}
	              key = ''
	              (1..15).each do |max_row|
	                @data = []
	                (0..3).each_with_index do |index, c_i|
	                  rrr = rr + max_row
	                  ccc = cc + c_i
	                  value = sheet_data.cell(rrr,ccc)
	                  if value.present?
	                    if (c_i == 0)
	                      key = value
	                      @block_hash[key] = {}
	                    else
	                      @block_hash[key][15*c_i] = value if key.present?
	                    end
	                    @data << value
	                  end
	                end
	                if @data.compact.reject { |c| c.blank? }.length == 0
	                  break # terminate the loop
	                end
	              end
	              @program.update(base_rate: @block_hash)
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def du_lp_llpas
  	@xlsx.sheets.each do |sheet|
      if (sheet == "DU & LP LLPAs")
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @subordinate_hash = {}
        @property_hash = {}
        @cashout_hash = {}
        primary_key = ''
        primary_key1 = ''
        secondary_key1 = ''
        secondary_key = ''
        ltv_key = ''
        ltv_key1 = ''
        cltv_key = ''
        new_key = ''
        # Adjustments
        (27..66).each do |r|
        	row = sheet_data.row(r)
        	@ltv_data = sheet_data.row(28)
        	@cltv_data = sheet_data.row(38)
        	if row.compact.count
        		(0..21).each do |cc|
        			value = sheet_data.cell(r,cc)
        			if value.present?
        				if value == "DU & LP LTV/FICO; Terms > 15 Years, Including ARMs"
									primary_key = "LoanType/Term/FICO/LTV"
									@adjustment_hash[primary_key] = {}
        				end
        				if value == "Subordinate Financing"
        					primary_key = "FinancingType/LTV/CLTV/FICO"
        					@subordinate_hash[primary_key] = {}
        				end
        				if value == "Multiple Unit Property"
        					primary_key = "PropertyType/LTV"
        					@property_hash[primary_key] = {}
        				end
        				if value == "Home Possible & Home Ready Adjustment Caps"
        					primary_key = "Caps/FICO/LTV"
        					@property_hash[primary_key] = {}
        				end
        				if value == "Cash Out"
        					primary_key1 = "RefinanceOption/FICO/LTV"
        					@cashout_hash[primary_key1] = {}
        				end
        				
        				# DU & LP LTV/FICO; Terms > 15 Years, Including ARMs
        				if r >= 29 && r <= 35 && cc == 3
        					secondary_key = get_value value
        					@adjustment_hash[primary_key][secondary_key] = {}
        				end
        				if r >= 29 && r <= 35 && cc >= 5 && cc <= 21
        					ltv_key = get_value @ltv_data[cc-1]
        					@adjustment_hash[primary_key][secondary_key][ltv_key] = {}
        					@adjustment_hash[primary_key][secondary_key][ltv_key] = value
        				end
        				# Subordinate Financing
        				if r >= 39 && r <= 42 && cc == 3
        					new_key = "DU"
        					@subordinate_hash[primary_key][new_key] = {}
        				end
        				if r >= 43 && r <= 46 && cc == 3
        					new_key = "LP"
        					@subordinate_hash[primary_key][new_key] = {}
        				end
        				if r >= 39 && r <= 42 && cc == 4 || r >= 43 && r <= 46 && cc == 4
        					secondary_key = get_value value
        					@subordinate_hash[primary_key][new_key][secondary_key]  = {}
        				end
        				if r >= 39 && r <= 42 && cc == 5 || r >= 43 && r <= 46 && cc == 5
        					cltv_key = get_value value
        					@subordinate_hash[primary_key][new_key][secondary_key][cltv_key] = {}
        				end
        				if r >= 39 && r <= 42 && cc >= 7 && cc <= 9 || r >= 43 && r <= 46 && cc >= 7 && cc <= 9
        					ltv_key = get_value @cltv_data[cc-1]
        					@subordinate_hash[primary_key][new_key][secondary_key][cltv_key][ltv_key] = {}
        					@subordinate_hash[primary_key][new_key][secondary_key][cltv_key][ltv_key] = value
        				end
        				if r == 47 && cc == 4
        					secondary_key = get_value value
        					@subordinate_hash[primary_key][secondary_key] = {}
        				end
        				if r == 47 && cc >= 7 && cc <= 9
        					ltv_key = get_value @cltv_data[cc-1]
        					@subordinate_hash[primary_key][secondary_key][ltv_key] = {}
        					@subordinate_hash[primary_key][secondary_key][ltv_key] = value
        				end
        				# Cash Out
        				if r >= 39 && r <= 45 && cc == 13
        					secondary_key1 = get_value value
        					@cashout_hash[primary_key1][secondary_key1] = {}
        				end
        				if r >= 39 && r <= 45 && cc >= 15 && cc <= 21
        					ltv_key1 = get_value @cltv_data[cc-1]
        					@cashout_hash[primary_key1][secondary_key1][ltv_key1] = {}
        					@cashout_hash[primary_key1][secondary_key1][ltv_key1] = value
        				end
        				if r >= 50 && r <= 51 && cc == 3
        					secondary_key = value.split("Property").first
        					@property_hash[primary_key][secondary_key] = {}
        					if @property_hash[primary_key][secondary_key] = {}
        						cc = cc + 6
        						new_value = sheet_data.cell(r,cc)
        						@property_hash[primary_key][secondary_key] = new_value
        					end
        				end
        				if r == 52 && cc == 3
        					secondary_key = value.split("Property").first
        					new_key = "0 < 80" 
        					@property_hash[primary_key][secondary_key] = {}
        					@property_hash[primary_key][secondary_key][new_key] = {}
        					if @property_hash[primary_key][secondary_key][new_key] = {}
        						cc = cc + 6
        						new_value = sheet_data.cell(r,cc)
        						@property_hash[primary_key][secondary_key][new_key] = new_value
        					end
        				end
        				if r == 53 && cc == 3
        					new_key = "80-85" 
        					@property_hash[primary_key][secondary_key][new_key] = {}
        					if @property_hash[primary_key][secondary_key][new_key] = {}
        						cc = cc + 6
        						new_value = sheet_data.cell(r,cc)
        						@property_hash[primary_key][secondary_key][new_key] = new_value
        					end
        				end
        				if r == 54 && cc == 3
        					new_key = "> 85" 
        					@property_hash[primary_key][secondary_key][new_key] = {}
        					if @property_hash[primary_key][secondary_key][new_key] = {}
        						cc = cc + 6
        						new_value = sheet_data.cell(r,cc)
        						@property_hash[primary_key][secondary_key][new_key] = new_value
        					end
        				end
        				if r >= 58 && r <= 60 && cc == 3
        					secondary_key = get_value value
        					@property_hash[primary_key][secondary_key] = {}
        				end
        				if r >= 58 && r <= 60 && cc == 5
        					ltv_key = get_value value
        					@property_hash[primary_key][secondary_key][ltv_key] = {}
        					if @property_hash[primary_key][secondary_key][ltv_key] = {}
        						cc = cc + 4
        						new_value = sheet_data.cell(r,cc)
        						@property_hash[primary_key][secondary_key][ltv_key] = new_value
        					end
        				end
        				if r == 65  && cc == 3
        					primary_key = value
        					@property_hash[primary_key] = {}
        					if @property_hash[primary_key] = {}
        						cc = cc + 6
        						new_value = sheet_data.cell(r,cc)
        						@property_hash[primary_key] = new_value
        					end
        				end
        				if r == 66  && cc == 3
        					primary_key = value
        					@property_hash[primary_key] = {}
        					if @property_hash[primary_key] = {}
        						cc = cc + 6
        						new_value = sheet_data.cell(r,cc)
        						@property_hash[primary_key] = new_value
        					end
        				end
        			end
        		end
        	end
        end
        adjustment = [@adjustment_hash,@subordinate_hash,@property_hash,@cashout_hash]
        make_adjust(adjustment,sheet)
        # create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def durp_lp_relief_llpas
  	@xlsx.sheets.each do |sheet|
      if (sheet == "DURP & LP Relief LLPAs")
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @subordinate_hash = {}
        @property_hash = {}
        @cashout_hash = {}
        primary_key = ''
        primary_key1 = ''
        secondary_key1 = ''
        secondary_key = ''
        ltv_key = ''
        ltv_key1 = ''
        cltv_key = ''
        new_key = ''
        # Adjustments
        (29..79).each do |r|
        	row = sheet_data.row(r)
        	@ltv_data = sheet_data.row(30)
        	@cltv_data = sheet_data.row(50)
        	if row.compact.count
        		(0..25).each do |cc|
        			value = sheet_data.cell(r,cc)
        			if value.present?
        				if value == "DURP LTV/FICO; Terms > 15 Years, Including ARMs"
									primary_key = "Durp/LoanType/Term/FICO/LTV"
									@adjustment_hash[primary_key] = {}
        				end
        				if value == "LP Relief LTV/FICO; Terms > 15 Years, Including ARMs"
									primary_key = "LP/LoanType/Term/FICO/LTV"
									@adjustment_hash[primary_key] = {}
        				end
        				if value == "Subordinate Financing"
        					primary_key = "FinancingType/LTV/CLTV/FICO"
        					@subordinate_hash[primary_key] = {}
        				end
        				if value == "Multiple Unit Property"
        					primary_key = "PropertyType/LTV"
        					@property_hash[primary_key] = {}
        				end
        				
        				# DURP LTV/FICO; Terms > 15 Years, Including ARMs
        				if r >= 31 && r <= 37 && cc == 3
        					secondary_key = get_value value
        					@adjustment_hash[primary_key][secondary_key] = {}
        				end
        				if r >= 31 && r <= 37 && cc >= 5 && cc <= 25
        					ltv_key = get_value @ltv_data[cc-1]
        					@adjustment_hash[primary_key][secondary_key][ltv_key] = {}
        					@adjustment_hash[primary_key][secondary_key][ltv_key] = value
        				end
        				# LP Relief LTV/FICO; Terms > 15 Years, Including ARMs
        				if r >= 41 && r <= 47 && cc == 3
        					secondary_key = get_value value
        					@adjustment_hash[primary_key][secondary_key] = {}
        				end
        				if r >= 41 && r <= 47 && cc >= 5 && cc <= 25
        					ltv_key = get_value @ltv_data[cc-1]
        					@adjustment_hash[primary_key][secondary_key][ltv_key] = {}
        					@adjustment_hash[primary_key][secondary_key][ltv_key] = value
        				end
        				# Subordinate Financing
        				if r >= 51 && r <= 53 && cc == 3
        					new_key = "DU"
        					@subordinate_hash[primary_key][new_key] = {}
        				end
        				if r >= 54 && r <= 59 && cc == 3
        					new_key = "LP"
        					@subordinate_hash[primary_key][new_key] = {}
        				end
        				if r >= 51 && r <= 53 && cc == 5 || r >= 54 && r <= 59 && cc == 5
        					secondary_key = get_value value
        					@subordinate_hash[primary_key][new_key][secondary_key]  = {}
        				end
        				if r >= 51 && r <= 53 && cc == 7 || r >= 54 && r <= 59 && cc == 7
        					cltv_key = get_value value
        					@subordinate_hash[primary_key][new_key][secondary_key][cltv_key] = {}
        				end
        				if r >= 51 && r <= 53 && cc >= 10 && cc <= 12 || r >= 54 && r <= 59 && cc >= 10 && cc <= 12
        					ltv_key = get_value @cltv_data[cc-1]
        					@subordinate_hash[primary_key][new_key][secondary_key][cltv_key][ltv_key] = {}
        					@subordinate_hash[primary_key][new_key][secondary_key][cltv_key][ltv_key] = value
        				end
        				if r == 60 && cc == 5
        					secondary_key = get_value value
        					@subordinate_hash[primary_key][secondary_key] = {}
        				end
        				if r == 60 && cc >= 7 && cc <= 12
        					ltv_key = get_value @cltv_data[cc-1]
        					@subordinate_hash[primary_key][secondary_key][ltv_key] = {}
        					@subordinate_hash[primary_key][secondary_key][ltv_key] = value
        				end
        				
        				# Multiple Unit Property
        				if r >= 63 && r <= 64 && cc == 3
        					secondary_key = value.split("Property").first
        					@property_hash[primary_key][secondary_key] = {}
        					if @property_hash[primary_key][secondary_key] = {}
        						cc = cc + 9
        						new_value = sheet_data.cell(r,cc)
        						@property_hash[primary_key][secondary_key] = new_value
        					end
        				end
        				if r == 65 && cc == 3
        					secondary_key = value.split("Property").first
        					new_key = "0 < 80" 
        					@property_hash[primary_key][secondary_key] = {}
        					@property_hash[primary_key][secondary_key][new_key] = {}
        					if @property_hash[primary_key][secondary_key][new_key] = {}
        						cc = cc + 9
        						new_value = sheet_data.cell(r,cc)
        						@property_hash[primary_key][secondary_key][new_key] = new_value
        					end
        				end
        				if r == 66 && cc == 3
        					new_key = "80-85" 
        					@property_hash[primary_key][secondary_key][new_key] = {}
        					if @property_hash[primary_key][secondary_key][new_key] = {}
        						cc = cc + 9
        						new_value = sheet_data.cell(r,cc)
        						@property_hash[primary_key][secondary_key][new_key] = new_value
        					end
        				end
        				if r == 67 && cc == 3
        					new_key = "> 85" 
        					@property_hash[primary_key][secondary_key][new_key] = {}
        					if @property_hash[primary_key][secondary_key][new_key] = {}
        						cc = cc + 9
        						new_value = sheet_data.cell(r,cc)
        						@property_hash[primary_key][secondary_key][new_key] = new_value
        					end
        				end
        			end
        		end
        	end
        end
        adjustment = [@adjustment_hash,@subordinate_hash,@property_hash,@cashout_hash]
        make_adjust(adjustment,sheet)
        # create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def programs
    @programs = @sheet_obj.programs
  end
  
  def single_program
  end

  def get_program
    @program = Program.find(params[:id])
  end

  private

  def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end

    def get_value value1
      if value1.present?
				if value1.include?("<=") || value1.include?("<")
          value1 = "0"+value1.tr('^0-9><%', '')
        elsif value1.include?(">=") || value1.include?(">")
        	value1 = value1.tr('^0-9><%', '')
        else
          value1
        end
      end
    end

    def make_adjust(block_hash, sheet)
      block_hash.each do |hash|
        Adjustment.create(data: hash,sheet_name: sheet)
      end
    end

    def read_sheet
      file = File.join(Rails.root,  'OB_Quicken_Loans3571.xls')
      @xlsx = Roo::Spreadsheet.open(file)
    end

    def program_property
      if @program.program_name.include?("30") || @program.program_name.include?("30/25 Year")
        term = 30
      elsif @program.program_name.include?("20")
        term = 20
      elsif @program.program_name.include?("15")
        term = 15
      elsif @program.program_name.include?("10 Year")
        term = 10
      elsif @program.program_name.include?("5 Year")
      	term = 5
      else
        term = nil
      end

      # Loan-Type
      if @program.program_name.include?("Fixed") || @program.program_name.include?("FIXED")
        loan_type = "Fixed"
      elsif @program.program_name.include?("ARM")
        loan_type = "ARM"
      elsif @program.program_name.include?("Floating")
        loan_type = "Floating"
      elsif @program.program_name.include?("Variable")
        loan_type = "Variable"
      else
        loan_type = nil
      end

      # Streamline Vha, Fha, Usda
      fha = false
      va = false
      usda = false
      streamline = false
      full_doc = false
      if @program.program_name.include?("FHA")
        streamline = true
        fha = true
        full_doc = true
      elsif @program.program_name.include?("VA")
        streamline = true
        va = true
        full_doc = true
      elsif @program.program_name.include?("USDA")
        streamline = true
        usda = true
        full_doc = true
      end

      # High Balance
      jumbo_high_balance = false
      if @program.program_name.include?("High Bal") || @program.program_name.include?("High Balance")
        jumbo_high_balance = true
      end

      # Arm Basic
      if @program.program_name.include?("3/1") || @program.program_name.include?("3 / 1")
        arm_basic = 3
      elsif @program.program_name.include?("5/1") || @program.program_name.include?("5 / 1")
        arm_basic = 5
      elsif @program.program_name.include?("7/1") || @program.program_name.include?("7 / 1")
        arm_basic = 7
      elsif @program.program_name.include?("10/1") || @program.program_name.include?("10 / 1")
        arm_basic = 10          
      end

      # Arm Advanced
      if @program.program_name.include?("2-2-5 ")
        arm_advanced = "2-2-5"
      end
      # Loan Limit Type
      if @program.program_name.include?("Non-Conforming")
        @program.loan_limit_type << "Non-Conforming"
      end
      if @program.program_name.include?("Conforming")
        @program.loan_limit_type << "Conforming"
      end
      if @program.program_name.include?("Jumbo")
        @program.loan_limit_type << "Jumbo"
      end
      if @program.program_name.include?("High Balance")
        @program.loan_limit_type << "High Balance"
      end
      @program.save
      @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline, jumbo_high_balance: jumbo_high_balance, arm_basic: arm_basic, arm_advanced: arm_advanced)
    end

    def create_program_association_with_adjustment(sheet)
      adjustment_list = Adjustment.where(sheet_name: sheet)
      program_list = Program.where(sheet_name: sheet)

      adjustment_list.each_with_index do |adj_ment, index|
        key_list = adj_ment.data.keys.first.split("/")
        program_filter1={}
        program_filter2={}
        include_in_input_values = false
        if key_list.present?
          key_list.each_with_index do |key_name, key_index|
            if (Program.column_names.include?(key_name.underscore))
              unless (Program.column_for_attribute(key_name.underscore).type.to_s == "boolean")
                program_filter1[key_name.underscore] = nil
              else
                if (Program.column_for_attribute(key_name.underscore).type.to_s == "boolean")
                  program_filter2[key_name.underscore] = true
                end
              end
            else
              if(Adjustment::INPUT_VALUES.include?(key_name))
                include_in_input_values = true
              end
            end
          end

          if (include_in_input_values)
            program_list1 = program_list.where.not(program_filter1)
            program_list2 = program_list1.where(program_filter2)

            if program_list2.present?
              program_list2.map{ |program| program.adjustments << adj_ment unless program.adjustments.include?(adj_ment) }
            end
          end
        end
      end
    end
end
