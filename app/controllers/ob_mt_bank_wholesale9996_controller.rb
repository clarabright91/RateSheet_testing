class ObMTBankWholesale9996Controller < ApplicationController
	before_action :get_sheet, only: [:programs, :rates]
  before_action :get_program, only: [:single_program]
  before_action :read_sheet, only: [:index,:rates]

  def index
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "Rates")
          @name = "M&T Bank National Wholesale"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
    end
  end

  def rates
  	@xlsx.sheets.each do |sheet|
      if (sheet == "Rates")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (5..722).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 1
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                # program_property sheet
                @block_hash = {}
                key = ''
                (1..15).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
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
                debugger
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_m_t_bank_wholesale9996_path(@sheet_obj)
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
  def read_sheet
    file = File.join(Rails.root,  'OB_M_&_T_Bank_Wholesale9996.xls')
    @xlsx = Roo::Spreadsheet.open(file)
  end
end