class Bank < ApplicationRecord
  has_many :programs

  SHEET_LINKS = []
	SHEET_LINKS << "https://www.loansifter.com/DownloadFile.aspx?RateSheetID=8570"
	SHEET_LINKS << "https://www.loansifter.com/DownloadFile.aspx?RateSheetID=10742"
	SHEET_LINKS << "https://www.loansifter.com/DownloadFile.aspx?RateSheetID=7575"
	SHEET_LINKS << "https://www.loansifter.com/DownloadFile.aspx?RateSheetID=11098"
	SHEET_LINKS << "https://www.loansifter.com/DownloadFile.aspx?RateSheetID=7019"
	SHEET_LINKS << "https://www.loansifter.com/DownloadFile.aspx?RateSheetID=3571"
	SHEET_LINKS << "https://www.loansifter.com/DownloadFile.aspx?RateSheetID=5907"
	SHEET_LINKS << "https://www.loansifter.com/DownloadFile.aspx?RateSheetID=4892"
	SHEET_LINKS << "https://www.loansifter.com/DownloadFile.aspx?RateSheetID=2982"
end
