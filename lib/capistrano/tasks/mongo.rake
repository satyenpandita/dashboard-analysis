namespace :mongo do
  %w(start stop restart).each do |task_name|
    desc "#{task} MongoDB"
    task task_name do
      on roles(:app), in: :sequence, wait: 5 do
        execute :sudo,  "service mongodb #{task_name}"
      end
    end
  end
end

