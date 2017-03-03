require "rvm/capistrano"
require "rvm/capistrano/alias_and_wrapp"

set :rvm_autolibs_flag, "read-only" # more info: rvm help autolibs
set :bundle_dir, ''
set :bundle_flags, '--system --quiet'
set :bundle_without,  [:development]

before 'deploy:setup', 'rvm:install_rvm'  # install/update RVM
before 'deploy:setup', 'rvm:install_ruby'

before 'deploy', 'rvm:create_alias'
before 'deploy', 'rvm:create_wrappers'

namespace :rvm do
  task :trust_rvmrc do
    run "rvm rvmrc trust #{release_path}"
  end
end

after "deploy", "rvm:trust_rvmrc"