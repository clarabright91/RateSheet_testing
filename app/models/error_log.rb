class ErrorLog < ApplicationRecord
  before_create :set_bank_name
  validate :validate_repeated_data
  def set_bank_name
    self.bank_name = Bank.joins(:sheets).where("sheets.name Like ?", self.sheet_name).first.name
  end

  def validate_repeated_data
    error_logs = self.class.where("sheet_name Like ? AND error_detail Like ? AND created_at > ?", self.sheet_name, self.error_detail, 24.hours.ago)
    if error_logs.any?
      errors.add(:created_at, 'Record is already exists')
    end
  end
end
