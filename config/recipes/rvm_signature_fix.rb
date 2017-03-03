namespace :rvm_signature_fix do
  desc "Install everything onto the server"
  task :install_signature do
    run "gpg --keyserver hkp://keys.gnupg.net --recv-keys 409B6B1796C275462A1703113804BB82D39DC0E3"
  end
end

after 'deploy:install', 'rvm_signature_fix:install_signature'