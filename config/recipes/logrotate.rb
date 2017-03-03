namespace :log do
  desc "Install log rotation script"
  task :rotate, :roles => :app do
    destination = "/etc/logrotate.d/#{application}"
    template "logrotate.erb", "/tmp/logrotate_#{application}"
    run "#{sudo} mv /tmp/logrotate_#{application} #{destination}"
    run "#{sudo} chown root #{destination}"
    run "#{sudo} chmod 600 #{destination}"
  end
  after "deploy:setup", "log:rotate"
end