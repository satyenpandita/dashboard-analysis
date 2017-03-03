namespace :upstart do
  desc 'Generate and upload Upstard configs for daemons needed by the app'
  task :update_configs, except: {no_release: true} do
    upstart_config_files = File.expand_path('../upstart/*.conf.erb', __FILE__)
    upstart_root         = '/etc/init'

    Dir[upstart_config_files].each do |upstart_config_file|
      config = ERB.new(IO.read(upstart_config_file)).result(binding)
      path   = "#{upstart_root}/#{File.basename upstart_config_file, '.erb'}"

      put config, path
    end
  end
end

after 'deploy:update_code', 'upstart:update_configs'