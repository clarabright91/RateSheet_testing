class ObHomePointFinancialWholesale11098Controller < ApplicationController
  before_action :get_sheet, only: [:programs, :conforming_standard, :conforming_high_balance, :fha_va_usda, :homestyle, :fha_203k, :durp, :lpoa, :err, :hlr, :homeready, :homepossible, :jumbo_select, :jumbo_choice]
  before_action :get_program, only: [:single_program, :program_property]

  def index
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "Conforming Standard")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "Home Point Financial Corporation"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def conforming_standard
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Conforming Standard")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (9..101).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 2))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              if r <= 53 || (r >= 80 && 11 <= 101)
                cc = 8*max_column + 1 # 1 / 9
              elsif r >= 55 && 11 <= 76
                cc = 5
              end
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "Margin 2.25%; Caps 2/2/5, index Libor"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end


  def conforming_high_balance
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Conforming High Balance")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (9..77).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 2))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 6*max_column + 2 # 2 /8
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "Margin 2.25%; Caps 2/2/5, index Libor"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..8).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def fha_va_usda
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FHA-VA-USDA")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (9..111).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 1 # 1 / 8 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..7).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end


  def homestyle
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "HomeStyle")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (10..54).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 2))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 12*max_column + 1 # 1 /13
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..8).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  # skip for same name programs
  def durp
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "DURP")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (9..128).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 2 #2/9/16
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  # skip for same name programs
  def lpoa
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "LPOA")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (11..169).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 8*max_column + 2 #2/10/18
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end


  def fha_203k
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FHA 203K")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (10..77).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 2))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 5
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..8).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def err
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "ERR")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (11..107).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 2 #2/9/16
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def hlr
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "HLR")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (11..107).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 2 #2/9/16
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def homeready
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "HomeReady")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (10..101).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              if r <= 77
                cc = 11*max_column + 5 # 5 / 16
              elsif r >= 80
                cc = 8*max_column + 3 # 3 / 11 / 19
              end
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def homepossible
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "HomePossible")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (9..78).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 1 #1/8/15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "Margin 2.25%; Caps 2/2/5" && @title != "Margin 2.25%; Caps 5/2/5"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def jumbo_select
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Select")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (10..54).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              if r <= 31
                cc = 10*max_column + 3 # 3 / 13
              elsif r >= 33
                cc = 7*max_column + 1 # 1 / 7 / 14
              end
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def jumbo_choice
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Choice")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (10..54).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              if r <= 31
                cc = 9*max_column + 3 # 3 / 12
              elsif r >= 33
                cc = 7*max_column + 1 # 1 / 8 / 15
              end
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def conforming_high_balance
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Conforming High Balance")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (9..77).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 2))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 6*max_column + 2 # 2 /8
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "Margin 2.25%; Caps 2/2/5, index Libor"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..8).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def fha_va_usda
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FHA-VA-USDA")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (9..111).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 1 # 1 / 8 / 15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..7).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end


  def homestyle
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "HomeStyle")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (10..54).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 2))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 12*max_column + 1 # 1 /13
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..8).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  # skip for same name programs
  # def durp
  #   @programs_ids = []
  #   file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
  #   xlsx = Roo::Spreadsheet.open(file)
  #   xlsx.sheets.each do |sheet|
  #     if (sheet == "DURP")
  #       sheet_data = xlsx.sheet(sheet)
  #       @programs_ids = []
  #       @block_hash = {}

  #       (11..107).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count >= 1) && (row.compact.count <= 3))
  #           rr = r + 1
  #           max_column_section = row.compact.count
  #           (0..max_column_section).each do |max_column|
  #             cc = 7*max_column + 2 #2/9/16
  #             @title = sheet_data.cell(r,cc)
  #             if @title.present?
  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               program_property @program
  #               @programs_ids << @program.id
  #               @program.adjustments.destroy_all
  #               @block_hash = {}
  #               key = ''
  #               (1..20).each do |max_row|
  #                 @data = []
  #                 (0..6).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if value.present?
  #                     if (c_i == 0)
  #                       key = value
  #                       @block_hash[key] = {}
  #                     else
  #                       @block_hash[key][15*c_i] = value
  #                     end
  #                     @data << value
  #                   end
  #                 end
  #                 if @data.compact.reject { |c| c.blank? }.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               @program.update(base_rate: @block_hash)
  #             end
  #           end
  #         end
  #       end
  #     end
  #   end
  #   redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  # end

  # skip for same name programs
  # def lpoa
  #   @programs_ids = []
  #   file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
  #   xlsx = Roo::Spreadsheet.open(file)
  #   xlsx.sheets.each do |sheet|
  #     if (sheet == "LPOA")
  #       sheet_data = xlsx.sheet(sheet)
  #       @programs_ids = []
  #       @block_hash = {}

  #       (11..107).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count >= 1) && (row.compact.count <= 3))
  #           rr = r + 1
  #           max_column_section = row.compact.count
  #           (0..max_column_section).each do |max_column|
  #             cc = 7*max_column + 2 #2/9/16
  #             @title = sheet_data.cell(r,cc)
  #             if @title.present?
  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               program_property @program
  #               @programs_ids << @program.id
  #               @program.adjustments.destroy_all
  #               @block_hash = {}
  #               key = ''
  #               (1..20).each do |max_row|
  #                 @data = []
  #                 (0..6).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if value.present?
  #                     if (c_i == 0)
  #                       key = value
  #                       @block_hash[key] = {}
  #                     else
  #                       @block_hash[key][15*c_i] = value
  #                     end
  #                     @data << value
  #                   end
  #                 end
  #                 if @data.compact.reject { |c| c.blank? }.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               @program.update(base_rate: @block_hash)
  #             end
  #           end
  #         end
  #       end
  #     end
  #   end
  #   redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  # end


  def fha_203k
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FHA 203K")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (10..77).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 2))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 5
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..8).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def err
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "ERR")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (11..107).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 2 #2/9/16
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def hlr
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "HLR")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (11..107).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 2 #2/9/16
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def homeready
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "HomeReady")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (10..101).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              if r <= 77
                cc = 11*max_column + 5 # 5 / 16
              elsif r >= 80
                cc = 8*max_column + 3 # 3 / 11 / 19
              end
              @title = sheet_data.cell(r,cc)
              begin
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def homepossible
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "HomePossible")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (9..78).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 1 #1/8/15
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "Margin 2.25%; Caps 2/2/5" && @title != "Margin 2.25%; Caps 5/2/5"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def jumbo_select
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Select")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (10..54).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              if r <= 31
                cc = 10*max_column + 3 # 3 / 13
              elsif r >= 33
                cc = 7*max_column + 1 # 1 / 7 / 14
              end
              begin
                @title = sheet_data.cell(r,cc)

                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def jumbo_choice
    @programs_ids = []
    file = File.join(Rails.root,  'OB_Home_Point_Financial_Wholesale11098.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Choice")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @block_hash = {}

        (10..54).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 3))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              if r <= 31
                cc = 9*max_column + 3 # 3 / 12
              elsif r >= 33
                cc = 7*max_column + 1 # 1 / 8 / 15
              end
              begin
                @title = sheet_data.cell(r,cc)

                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property @program
                  @programs_ids << @program.id
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..20).each do |max_row|
                    @data = []
                    (0..6).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_home_point_financial_wholesale11098_path(@sheet_obj)
  end

  def get_value value1
    if value1.present?
      if value1.include?("FICO <")
        value1 = "0"+value1.split("FICO").last
      elsif value1.include?("<")
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

  def programs
    @programs = @sheet_obj.programs
  end

  def single_program
  end

  private
    def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end

    def get_program
      @program = Program.find(params[:id])
    end

    def program_property value1
      # term
      if @program.program_name.include?("30 Year") || @program.program_name.include?("30Yr") || @program.program_name.include?("30 Yr") || @program.program_name.include?("30/25 Year")
        term = 30
      elsif @program.program_name.include?("20 Year")
        term = 20
      elsif @program.program_name.include?("15 Year")
        term = 15
      elsif @program.program_name.include?("10 Year")
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
        @program.loan_limit_type << "High-Balance"
      end
      @program.save
      @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline)
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



