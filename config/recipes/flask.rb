namespace :flask do
  desc "Setup Virtual Env"
  task :setup, roles: :web do
  	run "sudo chmod 777 #{shared_path}"
    run "cd #{shared_path}"
    run "sudo virtualenv --python=python3  #{shared_path}/env"
    run "source #{shared_path}/env/bin/activate"
    run "pip install -r #{current_path}/requirements.txt"
  end
  after "deploy:setup", "flask:setup"
end
