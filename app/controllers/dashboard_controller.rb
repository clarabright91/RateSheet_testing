class DashboardController < ApplicationController
  def index
    @banks = Bank.all
  end

  def calculator_index
  end
end
