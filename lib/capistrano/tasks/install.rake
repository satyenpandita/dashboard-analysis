namespace :deploy do
  desc "Install everything onto the server"
  task :install do
    on roles(:all) do |host|
        execute :sudo, "apt-get -y update"
        execute :sudo, "apt-get -y upgrade"
        execute :sudo, "apt-get -y install python3 python3-pip nginx gunicorn mongodb"
        execute "pip3 install virtualenv"
    end
  end
end
