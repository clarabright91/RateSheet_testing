Rails.application.routes.draw do

  root :to => "dashboard#index"
  # root :to => "import_files#index"
  resources :import_files, only: [:index] do
    member do
      get :programs
      get :import_government_sheet
      get :import_freddie_fixed_rate
      get :import_conforming_fixed_rate
      get :home_possible
      get :conforming_arms
      get :lp_open_acces_arms
      get :lp_open_access_105
      get :lp_open_access
      get :du_refi_plus_arms
      get :du_refi_plus_fixed_rate_105
      get :du_refi_plus_fixed_rate
      get :dream_big
      get :high_balance_extra
      get :freddie_arms
      get :jumbo_series_d
      get :jumbo_series_f
      get :jumbo_series_h
      get :jumbo_series_i
      get :jumbo_series_jqm
      get :import_HomeReadyhb_sheet
      get :import_homereddy_sheet
    end
  end
  # resources :ob_cmg_wholesales, only: [:index] do
  #   member do
  #     get :import_gov_sheet
  #   end
  # end
  # For details on the DSL available within this file, see http://guides.rubyonrails.org/routing.html
  resources :ob_cmg_wholesales, only: [:index] do
    member do
      get :programs
      get :gov
      get :agency
      get :durp
      get :oa
      get :jumbo_700
      get :jumbo_6200
      get :jumbo_7200_6700
      get :jumbo_6600
      get :jumbo_7600
      get :jumbo_6400
      get :jumbo_6800
      get :jumbo_6900_7900
      get :single_program
    end
  end

  resources :ob_cardinal_financial_wholesale10742, only: [:index] do
    member do
      get :ak
    end
  end

  resources :ob_allied_mortgage_group_wholesale8570, only: [:index] do
    member do
      get :programs
      get :fha
      get :va
      get :conf_fixed
      get :single_program
    end
  end

  resources :ob_newfi_wholesale7019, only: [:index] do
    member do
      get :programs
      get :biscayne_delegated_jumbo
      get :sequoia_portfolio_plus_products
      get :sequoia_expanded_products
      get :sequoia_investor_pro
      get :fha_buydown_fixed_rate_products
      get :fha_fixed_arm_products
      get :fannie_mae_homeready_products
      get :fnma_buydown_products
      get :fnma_conventional_fixed_rate
      get :fnma_conventional_high_balance
      get :fnma_conventional_arm
      get :olympic_piggyback_fixed
      get :olympic_piggyback_high_balance
      get :olympic_piggyback_arm
      get :single_program
    end
  end


  resources :ob_home_point_financial_wholesale11098, only: [:index] do
    member do
      get :programs
      get :conforming_standard
      get :conforming_high_balance
      get '/fha-va-usda' => 'ob_home_point_financial_wholesale11098#fha_va_usda'
      get :fha_203k
      get :homestyle
      get :durp
      get :lpoa
      get :err
      get :hlr
      get :homeready
      get :homepossible
      get :jumbo_select
      get :jumbo_choice
      get :single_program
    end
  end

  resources :ob_sun_west_wholesale_demo5907, only: [:index] do
    member do
      get :programs
      get :ratesheet
      get :single_program
    end
  end



  resources :ob_united_wholesale_mortgage4892, only: [:index] do
    member do
      get :programs
      get :single_program
      get :conv
      get :govt
      get :govt_arms
      get '/non-conf' => 'ob_united_wholesale_mortgage4892#non_conf'
      get :harp
    end
  end

  resources :ob_quicken_loans3571, only: [:index] do
    member do
      get :programs
      get :single_program
      get :ws_rate_sheet_summary
      get :ws_du_lp_pricing
      get :durp_lp_relief_pricing
      get :fha_usda_full_doc_pricing
      get :fha_streamline_pricing
      get :va_full_doc_pricing
      get :va_irrrl_pricing_govy_llpas
      get :na_jumbo_pricing_llpas
      get :du_lp_llpas
    end
  end

  resources :ob_union_home_mortgage_wholesale1711, only: [:index] do
    member do
      get :programs
      get :single_program
      get :conventional
    end
  end

  match "dashboard/index", to: 'dashboard#index', via: [:get, :post]
end
