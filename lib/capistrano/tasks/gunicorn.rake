namespace :gunicorn do
desc "Setup gunicorn configuration for this application"
  task :restart do
    on roles(:all) do |host|
        execute "cd #{current_path}"
        execute "kill -9 $(ps aux | grep -e gunicorn | awk '{ print $2 }') "
        execute "#{current_path}/env/bin/gunicorn app:app -b localhost:8000 --daemon"
        execute "supervisorctl restart celeryd"
    end
  end
end