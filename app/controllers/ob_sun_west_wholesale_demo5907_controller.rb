class ObSunWestWholesaleDemo5907Controller < ApplicationController
  before_action :get_sheet, only: [:programs, :ratesheet]
  before_action :get_program, only: [:single_program]
  def index
    file = File.join(Rails.root,  'OB_SunWest_Wholesale_Demo5907.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "RATESHEET")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "SunWest Wholesale"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def ratesheet
    file = File.join(Rails.root,  'OB_SunWest_Wholesale_Demo5907.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "RATESHEET")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        first_key = []
        @key_data = []
        @conf_adjustment = {}
        k_value = ''
        value1 = ''
        range1 = 374
        range2 = 404
        range1_a = 782
        range2_b = 799

        # # Agency Conforming Programs
        # (156..320).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 4))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 5*max_column + 2 # 2 / 7 / 12 / 17
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? && @title != "Rate"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..4).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end

        #PRICE ADJUSTMENTS: CONFORMING PROGRAMS //adjustment
        # (range1..range2).each do |r|
        #   (0..sheet_data.last_column).each do |cc|
        #     value = sheet_data.cell(r,cc)
        #     # if value == "LOAN TERM > 15 YEARS"
        #     #   first_row = 377
        #     #   end_row = 384
        #     #   last_column = 10
        #     #   first_column = 2
        #     #   ltv_row = 375
        #     #   ltv_adjustment range1, range2, sheet_data, first_row, end_row,sheet,first_column, last_column, ltv_row
        #     # end

        #     # if value == "CASH OUT REFINANCE "
        #     #   first_row = 389
        #     #   end_row = 396
        #     #   first_column = 2
        #     #   last_column = 6
        #     #   ltv_row = 387
        #     #   ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row
        #     # end

        #     # if value == "ADDITIONAL LPMI ADJUSTMENTS"
        #     #   first_row = 390
        #     #   end_row = 393
        #     #   first_column = 9
        #     #   last_column = 12
        #     #   ltv_row = 388
        #     #   ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row
        #     # end

        #     # if value == "LPMI COVERAGE BASED ADJUSTMENTS"
        #     #   first_row = 399
        #     #   end_row = 404
        #     #   first_column = 9
        #     #   last_column = 12
        #     #   ltv_row = 397
        #     #   ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row
        #     # end

        #     # if value == "LPMI COVERAGE BASED ADJUSTMENTS"
        #     #   first_row = 399
        #     #   end_row = 404
        #     #   first_column = 9
        #     #   last_column = 12
        #     #   ltv_row = 397
        #     #   ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row
        #     # end

        #     # if value == "SUBORDINATE FINANCING" #not completed 2 more remaining adjustmest
        #     #   first_row = 400
        #     #   end_row = 404
        #     #   first_column = 2
        #     #   last_column = 7
        #     #   ltv_row = 399
        #     #   ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row
        #     # end
        #   end
        # end

        # # FHLMC HOME Programs
        # (708..760).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 4))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 5*max_column + 2 # 2 / 7 / 12 / 17
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? && @title != "Rate"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..4).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end

        ## PRICE ADJUSTMENTS: FHLMC HOME POSSIBLE / HOMEONE / SUPER CONFORMING //adjustment
        (range1_a..range2_b).each do |r|
          (0..sheet_data.last_column).each do |cc|
            value = sheet_data.cell(r,cc)

            if value == "LOAN TERM > 15 YEARS"
              first_row = 400
              end_row = 404
              first_column = 2
              last_column = 7
              ltv_row = 399
              ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row
            end
          end
        end

        # #Non-Confirming: Sigma Programs
        # (1101..1179).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 4))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 5*max_column + 2 # 2 / 7 / 12 / 17
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? && @title != "ARM INFORMATION"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..4).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end

        # #Non-Confirming: JW
        # (1386..1547).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 4))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 5*max_column + 2 # 2 / 7 / 12 / 17
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? && @title != "Rate"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..4).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end

        # #Non-Confirming: Government Not Completed
        # (2180..2278).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 4))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 5*max_column + 2 # 2 / 7 / 12 / 17
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? #&& @title != "PROGRAM SPECIFIC PRICE ADJUSTMENTS"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..4).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*c_i
        #                 @program.save
        #               end
        #               # debugger if r > 2254 && cc > 7
        #               @block_hash[main_key][key][15*c_i] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end

        # #NON-QM: SIGMA NO CREDIT EVENT PLUS Program // issue not done
        # (2624..2675).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8 / 11 / 14
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? && @title != "ARM INFORMATION"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..2).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end

        # # NON-QM: SIGMA SEASONED CREDIT EVENT, SIGMA RECENT CREDIT EVENT Programs ankit
        # (2434..2477).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1))
        #     rr = r + 1
        #     (0..9).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8 / 11 / 14
        #       @title = sheet_data.cell(r,cc)

        #       if @title.present? && @title != "Rate" && cc <= 8 && @title.class != Float
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..2).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end


        # # NON-QM: R.E.A.L PRIME ADVANTAGE Programs done
        # (3237..3249).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 4))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8 / 11 / 14
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? && @title != "MAXIMUM PRICE"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..2).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end

        # #NON-QM: R.E.A.L CREDIT ADVANTAGE - A Program // issue not done
        # (2936..2948).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8 / 11 / 14
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? && @title != "Rate" && @title != "MAXIMUM PRICE" && @title != "LOAN AMOUNT"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..2).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       debugger
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end

        # #NON-QM: R.E.A.L CREDIT ADVANTAGE - B, B-, C Program 15 Year Fixed error
        # (3089..3101).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8 / 11 / 14
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? && @title != "Rate" && @title != "MAXIMUM PRICE" && @title != "LOAN AMOUNT"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..2).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end

        # #NNON-QM: R.E.A.L INVESTOR INCOME - A Program
        # (3237..3249).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 4))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8 / 11 / 14
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? && @title != "MAXIMUM PRICE"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..2).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end

        # #NON-QM: R.E.A.L INVESTOR INCOME - B, B- Program
        # (3334..3346).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8 / 11 / 14
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? && @title != "MAXIMUM PRICE"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..2).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end

        #  # NON-QM: R.E.A.L DSC RATIO Programs
        # (3433..3445).each do |r|
        #   row = sheet_data.row(r)
        #   if ((row.compact.count >= 1) && (row.compact.count <= 7))
        #     rr = r + 1
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 3*max_column + 2 # 2 / 5 / 8 / 11 / 14
        #       @title = sheet_data.cell(r,cc)
        #       if @title.present? && @title != "LOAN AMOUNT"
        #         @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #         program_property @program
        #         @programs_ids << @program.id
        #       end

        #       @program.adjustments.destroy_all
        #       @block_hash = {}
        #       key = ''
        #       if @program.term.present?
        #         main_key = "Term/LoanType/InterestRate/LockPeriod"
        #       else
        #         main_key = "InterestRate/LockPeriod"
        #       end
        #       @block_hash[main_key] = {}
        #       (1..50).each do |max_row|
        #         @data = []
        #         (0..2).each_with_index do |index, c_i|
        #           rrr = rr + max_row
        #           ccc = cc + c_i
        #           value = sheet_data.cell(rrr,ccc)
        #           if value.present?
        #             if (c_i == 0)
        #               key = value
        #               @block_hash[main_key][key] = {}
        #             else
        #               if @program.lock_period.length <= 3
        #                 @program.lock_period << 15*(c_i+1)
        #                 @program.save
        #               end
        #               @block_hash[main_key][key][15*(c_i+1)] = value
        #             end
        #             @data << value
        #           end
        #         end
        #         if @data.compact.reject { |c| c.blank? }.length == 0
        #           break # terminate the loop
        #         end
        #       end
        #       if @block_hash.values.first.keys.first.nil? || @block_hash.values.first.keys.first == "Rate"
        #         @block_hash.values.first.shift
        #       end
        #       @program.update(base_rate: @block_hash)
        #     end
        #   end
        # end
      end
    end
    redirect_to programs_ob_sun_west_wholesale_demo5907_path(@sheet_obj)
  end

  def programs
    @programs = @sheet_obj.programs
  end

  def single_program
  end

  private

    def get_value value1
      if value1.present?
        if value1.include?("FICO <")
          value1 = "0"+value1.split("FICO").last
        elsif value1.include?("<=") || value1.include?(">=")
          value1 = "0"+value1
        elsif value1.include?("FICO")
          value1 = value1.split("FICO ").last.first(9)
        elsif value1 == "Investment Property"
          value1 = "Property/Type"
        else
          value1
        end
      end
    end

    def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end

    def get_program
      @program = Program.find(params[:id])
    end

    def program_property value1
      # term
      if @program.program_name.include?("30 Year") || @program.program_name.include?("30Yr") || @program.program_name.include?("30 Yr") || @program.program_name.include?("30/25 Year") || @program.program_name.include?("30 YR")
        term = 30
      elsif @program.program_name.include?("20 Year") || @program.program_name.include?("20 YR")
        term = 20
      elsif @program.program_name.include?("15 Year") || @program.program_name.include?("15 YR")
        term = 15
      elsif @program.program_name.include?("10 Year") || @program.program_name.include?("10 YR")
        term = 10
      else
        term = nil
      end

      # Loan-Type
      if @program.program_name.include?("Fixed")
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
      if @program.program_name.include?("High Bal")
        @jumbo_high_balance = true
      end

       # Program Category
      if @program.program_name.include?("F30/F25")
        @program_category = "F30/F25"
      elsif @program.program_name.include?("F15")
        @program_category = "F15"
      elsif @program.program_name.include?("f30J")
        @program_category = "f30J"
      elsif @program.program_name.include?("F15S")
        @program_category = "F15S"
      elsif @program.program_name.include?("F30JS")
        @program_category = "F30JS"
      elsif @program.program_name.include?("F5YT")
        @program_category = "F5YT"
      elsif @program.program_name.include?("F5YTS")
        @program_category = "F5YTS"
      elsif @program.program_name.include?("F5YTJ")
        @program_category = "F5YTJ"
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
      @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline)
    end

    def ltv_adjustment range1, range2, sheet_data, first_row, end_row, sheet, first_column, last_column, ltv_row
     @adjustment_hash = {}
     primary_key = ''
     ltv_key = ''
     cltv_key = ''
     (range1..range2).each do |r|
       row = sheet_data.row(r)
       @ltv_data = sheet_data.row(ltv_row)
       if row.compact.count >= 1
         (0..last_column).each do |cc|
           value = sheet_data.cell(r,cc)
           if value.present?
             if value == "LOAN TERM > 15 YEARS" #"CASH OUT REFINANCE " "ADDITIONAL LPMI ADJUSTMENTS" "LPMI COVERAGE BASED ADJUSTMENTS"
               primary_key = "LoanType/Term/LTV/FICO"
               @adjustment_hash[primary_key] = {}
             end
             if r >= first_row && r <= end_row && cc == first_column
               ltv_key = value
               @adjustment_hash[primary_key][ltv_key] = {}
             end
             if r >= first_row && r <= end_row && cc > first_column && cc <= last_column
               cltv_key = get_value @ltv_data[cc-2]
               @adjustment_hash[primary_key][ltv_key][cltv_key] = {}
               @adjustment_hash[primary_key][ltv_key][cltv_key] = value
             end
           end
         end
       end
     end
     adjustment = [@adjustment_hash]
     make_adjust(adjustment,sheet)
    end

    def make_adjust(block_hash, sheet)
      block_hash.each do |hash|
        Adjustment.create(data: hash,sheet_name: sheet)
      end
    end
end
