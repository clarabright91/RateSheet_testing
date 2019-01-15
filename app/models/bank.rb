class Bank < ApplicationRecord
  has_many :programs
  has_many :sheets
end
