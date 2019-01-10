Rails.application.routes.draw do
  get 'ob_cmg_wholesales/index'
  get 'ob_cmg_wholesales/import_gov_sheet'
  get 'ob_cmg_wholesales/import_agency_sheet'
  get 'ob_cmg_wholesales/import_durp_sheet'
  get 'ob_cmg_wholesales/import_oa_sheet'
  get 'ob_cmg_wholesales/import_jumbo700_sheet'
  get 'ob_cmg_wholesales/import_jumbo6200_sheet'
  get 'ob_cmg_wholesales/import_jumbo7200_6700_sheet'
  get 'ob_cmg_wholesales/import_jummbo6600_sheet'
  get 'ob_cmg_wholesales/import_jummbo7600_sheet'
  get 'ob_cmg_wholesales/import_jummbo6400_sheet'
  get 'ob_cmg_wholesales/import_jummbo6800_sheet'
  get 'ob_cmg_wholesales/import_jumbo6900_7900_sheet'
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
  resources :ob_cmg_wholesales do
    member do
      get :programs
    end
  end

  match "dashboard/index" ,via: [:get, :post]
  
  end
