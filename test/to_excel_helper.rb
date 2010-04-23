begin
  require 'rubygems'
  require 'test/unit'
  require 'active_support'
  require 'to_excel'
end

class User
  COLUMNS = %w(id name age)
  
  attr_accessor *COLUMNS

  def self.human_attribute_name(attribute)
    attribute.to_s.humanize
  end

  def initialize(params={})
    params.each { |key, value| self.send("#{key}=", value); }
    self
  end

  def attributes
    COLUMNS.inject({}) { |attributes, attribute| attributes.merge(attribute => send(attribute)) }
  end

  def is_old?
    age > 40
  end
end

class Computer
  COLUMNS = %w(id brand portable user)
  
  attr_accessor *COLUMNS

  def self.human_attribute_name(attribute)
    attribute.to_s.humanize
  end

  def initialize(params={})
    params.each { |key, value| self.send("#{key}=", value); }
    self
  end

  def attributes
    COLUMNS.inject({}) { |attributes, attribute| attributes.merge(attribute => send(attribute)) }
  end

end
