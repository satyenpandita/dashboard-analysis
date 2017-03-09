namespace :deploy do
  desc "Install new Packages"
  task :setup do
    on roles(:all) do |host|
        execute "source #{shared_path}/env/bin/activate"
    end
  end
end