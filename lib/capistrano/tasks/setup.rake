namespace :deploy do
  desc "Install new Packages"
  task :setup do
    on roles(:all) do |host|
        execute :sudo, "/usr/bin/easy_install virtualenv"
        execute "virtualenv --python=python3 #{current_path}/env"
        execute "source #{current_path}/env/bin/activate"
        execute "#{current_path}/env/bin/pip3 install -r #{current_path}/requirements.txt"
        execute "#{current_path}/env/bin/pip3 install gunicorn"
    end
  end

  task :refresh do
    on roles(:all) do |host|
        execute "virtualenv --python=python3 #{current_path}/env"
        execute "source #{current_path}/env/bin/activate"
        execute "source #{current_path}/env/bin/activate"
        execute "#{current_path}/env/bin/pip3 install -r #{current_path}/requirements.txt"
        execute "#{current_path}/env/bin/pip3 install gunicorn"
        execute :sudo, "chmod 777 -R /var/www/auroville/releases"
    end
  end

end