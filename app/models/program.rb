class Program < ApplicationRecord
  belongs_to :bank

  def get_adjustments
    Adjustment.where(sheet_name: self.sheet_name)
  end
end
