namespace :deploy do
  desc "Setup Virtual Env"
  task :setup do
    on roles(:all) do |host|
        execute "source #{shared_path}/env/bin/activate"
        execute "pip3 install -r #{current_path}/requirements.txt"
    end
  end
end
