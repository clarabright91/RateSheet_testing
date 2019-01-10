class Program < ApplicationRecord
  belongs_to :bank, optional: true
  has_many :program_adjustments
  has_many :adjustments, through: :program_adjustments

  def get_adjustments
    Adjustment.where(sheet_name: self.sheet_name)
  end
end
