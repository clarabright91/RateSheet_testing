class Sheet < ApplicationRecord
  belongs_to :bank
  has_many :programs
end
