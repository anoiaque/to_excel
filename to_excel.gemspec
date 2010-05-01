spec = Gem::Specification.new do |s| 
  s.name = "to_excel"
  s.version = "1.0"
  s.author = "Philippe Cantin"
  s.email = ""
  s.homepage = "http://github.com/anoiaque/to_excel"
  s.platform = Gem::Platform::RUBY
  s.summary = "Export ruby objects  to excel file. Allow many collections of different class. Can include associations attributes."
  s.files = FileList["{lib}/*.rb"].to_a
  s.require_path = "lib"
  s.test_files = FileList["{test}/*test.rb"].to_a
  s.has_rdoc = false
  s.extra_rdoc_files = ["README"]
end
 
Rake::GemPackageTask.new(spec) do |pkg| 
  pkg.need_tar = false 
end 