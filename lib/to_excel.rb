class Array
  def to_excel(options = {})
    output = '<?xml version="1.0" encoding="UTF-8"?><Workbook xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40" xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office"><Worksheet ss:Name="Sheet1"><Table>'

    if self.first.is_a?(Array) && self.first.any?
      unless options[:headers] == false
        output << "<Row>"

        if options[:headers].is_a? Array
          options[:headers].each do |title|
            output << "<Cell><Data ss:Type=\"String\">#{title}</Data></Cell>"
          end  
        else  
          klasses = self.map(&:first).map(&:class)
          klasses.each do |klass|
            options[:map][klass].each do |attribute|
             output << "<Cell><Data ss:Type=\"String\">#{klass.human_attribute_name(attribute)}</Data></Cell>"
            end 
          end
        end

        output << "</Row>"
      end
      
      (0..self.map(&:length).max - 1).each do |n|
        items = self.map {|collection| collection[n] || collection.first.class.new}
        output << "<Row>"
        items.each do |item|
          options[:map][item.class].each do |column|
            value = column.is_a?(Array) ? column.inject(item){|result,column| result.send(column) if result} : item.send(column)
            output << "<Cell><Data ss:Type=\"#{value.is_a?(Integer) ? 'Number' : 'String'}\">#{value}</Data></Cell>"
          end
        end
        output << "</Row>"
      end
   end

   output << '</Table></Worksheet></Workbook>'
  end
end