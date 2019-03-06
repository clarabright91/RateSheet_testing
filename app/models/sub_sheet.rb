class SubSheet < ApplicationRecord
  belongs_to :sheet
  has_many :programs
end
