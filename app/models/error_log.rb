class ErrorLog < ApplicationRecord
  before_create :set_bank_name

  def set_bank_name
    self.bank_name = Bank.joins(:sheets).where("sheets.name Like ?", self.sheet_name).first.name
  end
end
