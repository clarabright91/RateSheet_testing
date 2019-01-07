class DashboardController < ApplicationController
  def index
    @banks = Bank.all
  end
end
