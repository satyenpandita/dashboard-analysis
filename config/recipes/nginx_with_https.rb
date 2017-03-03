namespace :nginx do
  desc "Install latest stable release of nginx"
  task :install, roles: :web do
    run "#{sudo} apt-get -y update"
    run "#{sudo} apt-get -y install nginx"
  end
  after "deploy:install", "nginx:install"

  desc "Setup nginx configuration for this application"
  task :setup, roles: :web do
    template "nginx_unicorn_https.erb", "/tmp/nginx_conf"
    run "#{sudo} mv /tmp/nginx_conf /etc/nginx/sites-enabled/#{application}"
    run "#{sudo} rm -f /etc/nginx/sites-enabled/default"
    restart
  end
  after "deploy:setup", "nginx:setup"

  desc "Setup ssl for nginx"
  task :setup_ssl, roles: :web do
    template "ssl_certificate.erb", "/tmp/certificate"
    template "ssl_key.erb", "/tmp/key"
    run "#{sudo} mkdir -p /etc/nginx/ssl"
    run "#{sudo} mv /tmp/certificate /etc/nginx/ssl/#{application}.crt"
    run "#{sudo} mv /tmp/key /etc/nginx/ssl/#{application}.key"
    restart
  end
  before "nginx:setup", "nginx:setup_ssl"
  
  %w[start stop restart].each do |command|
    desc "#{command} nginx"
    task command, roles: :web do
      run "#{sudo} service nginx #{command}"
    end
  end
end
