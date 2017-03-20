namespace :deploy do
  desc "Install new Packages"
  task :setup do
    on roles(:all) do |host|
        execute :sudo, "/usr/bin/easy_install virtualenv"
        execute "virtualenv --python=python3 #{current_path}/env"
        execute "source #{current_path}/env/bin/activate"
        execute "pip3 install -r #{current_path}/requirements.txt"
        execute "pip3 install gunicorn"
    end
  end
end