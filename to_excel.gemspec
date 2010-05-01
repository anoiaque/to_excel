specifications = Gem::Specification.new do |spec| 
  spec.name = "to_excel"
  spec.version = "1.0"
  spec.author = "Philippe Cantin"
  spec.homepage = "http://github.com/anoiaque/to_excel"
  spec.platform = Gem::Platform::RUBY
  spec.summary = "Export ruby objects  to excel file. Allow many collections of different classpec. Can include associations attributespec."
  spec.description = "Export ruby objects  to excel file"
  spec.files = Dir['lib/**/*.rb']
  spec.require_path = "lib"
  spec.test_files  = Dir['test/**/*.rb']
  spec.has_rdoc = false
  spec.extra_rdoc_files = ["README.rdoc"]
end