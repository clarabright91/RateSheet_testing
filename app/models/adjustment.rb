class Adjustment < ApplicationRecord
  #validates :program_title, presence: true, uniqueness: {message: "already exists"}
  before_create :check_details
  has_many :program_adjustments
  has_many :programs, through: :program_adjustments
  def check_details
  end

  MAIN_KEYS = {
    "All Conforming ARMs (Does not include LP Open Access)" => "Conforming/Term/LTV/FICO",
    "All Fixed Conforming\n(does not apply to terms <=15yrs)" => "Conforming/RateType/Term/LTV/FICO",
    "All Fixed Conforming\n(does not apply to terms <=15yrs with LTV <=95)" => "Conforming/RateType/Term/LTV/FICO",
    "All Conforming\n(does not apply to Fixed terms <=15yrs with LTV <=95)" => "Conforming/RateType/Term/LTV/FICO",
    "Cash-Out" => "RefinanceOption/LTV/FICO",
    "Lender Paid MI Adjustments" => "Term/LTV/FICO",
    "Premium Adjustments" => "LPMI/PremiumType/FICO",
    "LTV Adjustments" => "LPMI/Term/LTV/FICO",
    "Loan Size Adjusters" => "LoanAmount/LoanPurpose",
    "Number Of Units" => "PropertyType/LTV",
    "Subordinate Financing" => "FinancingType/LTV/CLTV/FICO",
    "Misc Adjusters" => "PropertyType/LTV/Term",
    "Non Owner Occupied" => "PropertyType/LTV",
    "Loan Size Adjustments" => "LoanType/Conforming/LTV/FICO",
    "Super Conforming" => "Conforming/LTV/FICO",
    "Super Conforming Adjustments" => "LoanPurpose/RefinanceOption",
    "Subordinate Financing\nExcludes Community Seconds®" => "FinancingType/LTV/CLTV/FICO"
  }

  DREAM_BIG_ADJUSTMENT = {
    4 => {"<=50" => "0"},
    5 => {"50.01 - 55" => "50.01"},
    6 => {"55.01 - 60" => "55.01"},
    7 => {"60.01 - 65" => "60.01"},
    9 => {"65.01 - 70" => "65.01"},
    10 => {"70.01 - 75" => "70.01"},
    11 => {"75.01 - 80" => "75.01"},
    12 => {"80.01 - 85" => "80.01"},
    14 => {"85.01 - 90" => "85.01"},
    rows: {
      40 => {"680 - 699" => "680"},
      41 => {"700 - 719" => "700"},
      42 => {"720 - 739" => "720"},
      43 => {"740 - 759" => "740"},
      44 => {"760-779" => "760"},
      45 => {">=780"=> "780"},
      46 => {"Purchase" => "Purchase"},
      47 => {"Cash Out Refinance" => "Cash Out Refinance"},
      48 => {"Rate & Term Refinance" => "Rate & Term Refinance"},
      50 => {"Non Owner Occupied" => "Non Owner Occupied"},
      51 => {"> 80 LTV No MI" => "> 80 LTV No MI"},
      55 => {"680 - 699" => "680"},
      56 => {"700 - 719" => "700"},
      57 => {"720 - 739" => "720"},
      58 => {"740 - 759" => "740"},
      59 => {"760-779" => "760"},
      60 => {">=780"=> "780"},
      61 => {"Purchase" => "Purchase"},
      62 => {"Cash Out Refinance" => "Cash Out Refinance"},
    },
    arm_column: {
      4 => {"<=50" => "0"},
      5 => {"50.01 - 55" => "50.01"},
      6 => {"55.01 - 60" => "55.01"},
      8 => {"60.01 - 65" => "60.01"},
      9 => {"65.01 - 70" => "65.01"},
      10 => {"70.01 - 75" => "70.01"},
      11 => {"75.01 - 80" => "75.01"},
      12 => {"80.01 - 85" => "80.01"},
      14 => {"85.01 - 90" => "85.01"},
    }
  }

  JUMBO_SERIES_I_ADJUSTMENT = {
    5 => {"≤ 60" => "0"},
    6 => {"60.01-65" => "60.01"},
    7 => {"65.01-70" => "65.01"},
    8 => {"70.01-75" => "70.01"},
    10 => {"75.01-80" => "75.01"},
    14 => {"≤ 60" => "0"},
    16 => {"60.01-65" => "60.01"},
    17 => {"65.01-70" => "65.01"},
    18 => {"70.01-75" => "70.01"},
    19 => {"75.01-80" => "75.01"},
    rows: {
      41 => {"< 700" => "0"},
      42 => {"740-759" => "740"},
      43 => {"720-739" => "720"},
      44 => {"700-719" => "700"},
      45 => {"680-699" => "680"},
      50 => {"≤ $1MM" => "0"},
      51 => {"$1MM - $1.5MM" => "$1MM"},
      52 => {"$1.5MM - $2MM" => "$1.5MM"},
      53 => {"$2MM - $2.5MM" => "$2MM"},
      58 => {"2nd Home" => "2nd Home"},
      59 => {"Purchase (15 Yr Fixed ONLY)" => "Purchase (15 Yr Fixed ONLY)"},
      60 => {"C/O Refinance" => "C/O Refinance"},
      61 => {"2-4 Unit" => "2-4 Unit"},
      62 => {"DTI > 40%" => "DTI > 40%"}
    }
  }

  HIGH_BALANCE_ADJUSTMENT = {
    4 => {"<= 60" => "0"},
    5 => {"60.01 - 70" => "60.01"},
    6 => {"70.01 - 75" => "70.01"},
    7 => {"75.01 - 80" => "75.01"},
    8 => {"80.01 - 85" => "80.01"},
    9 => {"85.01 - 90" => "85"},
    rows: {
      28 => {">=760" => "760"},
      29 => {"740-759" => "740"},
      30 => {"720-739" => "720"},
      31 => {"700-719" => "700"},
      32 => {"680-699" => "680"},
      34 => {">=760" => "760"},
      35 => {"740-759" => "740"},
      36 => {"720-739" => "720"},
      37 => {"700-719" => "700"},
      38 => {"680-699" => "680"},
    },
    subordinate: {
      4 => {"< 720" => "0"},
      5 => {">= 720" => "720"}
    }
  }

  ALL_IP = {
    5 => {"<=80" => "0"},
    6 => {"80.01 - 85" => "80.01"},
    7 => {"> 85" => "85"},
    9 => {"< 720" => "0"},
    10 => {">= 720" => "720"},
    11 => {"< 620" => "0"},
    12 => {"620 - 639" => "620"},
    13 => {"640 - 659" => "640"},
    14 => {"660 - 679" => "660"},
    16 => {"680 - 699" => "680"},
    17 => {"700 - 719" => "700"},
    18 => {"720 - 739" => "720"},
    19 => {">= 740" => "740"},
    rows: {
      40 => {"<= 60" => "0"},
      41 => {"60.01 - 70" => "60.01"},
      42 => {"70.01 - 75" => "70.01"},
      43 => {"75.01 - 80" => "75.01"},
      44 => {"80.01 - 85" => "80.01"},
      45 => {"> 85 "=> "85"},
      48 => {"<=75" => "0"},
      49 => {"<=65" => "0"},
      50 => {"65.01-75" => "65.01"},
      51 => {"75.01-80" => "75.01"},
      52 => {"80.01-90" => "80.01"},
      53 => {"90.01-95" => "90.01"},
      54 => {"All" => "All"},
      57 => {"2 Units" => "2 Unit"},
      58 => {"3-4 units" => "3-4 Unit"},
      61 => {"<$50,000" => "0"},
      62 => {"$50,000 - $99,999" => "$50,000"},
      63 => {"$100,000 - $149,999" => "$100,000"},
      64 => {"$150,000 - $199,999" => "$150,000"},
      65 => {"$200,000 - $249,999" => "$200,000"},
      66 => {"$250,000 - $299,999" => "$250,000"},
      67 => {"$300,000 - Conforming Limit" => "$300,000"},
    },
    cltv: {
      48 => {"<=80" => "0"},
      49 => {"80.01 - 95" => "80.01"},
      50 => {"80.01 - 95" => "80.01"},
      51 => {"76.01 - 95" => "76.01"},
      52 => {"81.01 - 95" => "81.01"},
      53 => {"91.01 - 95" => "91.01"},
      54 => {"> 95" => "95"}
    }
  }

  INPUT_VALUES = [
    "PropertyType",
    "FinancingType",
    "PremiumType",
    "LTV",
    "FICO",
    "RefinanceOption",
    "MiscAdjuster",
    "LPMI",
    "Coverage",
    "LoanAmount",
    "CLTV",
    "DTI",
    "InterestRate",
    "LockPeriod",
    "State",
    "Term",
    "LoanType"
  ]
end
