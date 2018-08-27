def template(from, to)
  full_path = File.expand_path("../templates/#{from}", __FILE__)
  erb = StringIO.new(ERB.new(File.read(full_path)).result(binding))
  upload! erb, to
end