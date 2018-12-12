Rails.application.routes.draw do
  root :to => "import_files#index"
  resources :import_files, only: [:index] do
    member do
      get :programs
      get :import_government_sheet
    end
  end
  # For details on the DSL available within this file, see http://guides.rubyonrails.org/routing.html
end
