Rails.application.routes.draw do
  root :to => "import_files#index"
  resources :import_files, only: [:index] do
    member do
      get :programs
      get :import_government_sheet
      get :import_freddie_fixed_rate
      get :import_conforming_fixed_rate
      get :home_possible
      get :conforming_arms
    end
  end
  # For details on the DSL available within this file, see http://guides.rubyonrails.org/routing.html
end
