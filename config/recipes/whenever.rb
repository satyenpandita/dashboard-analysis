set :whenever_environment, defer { stage }
set :whenever_command, "bundle exec whenever"
require "whenever/capistrano"