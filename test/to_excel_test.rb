require  'to_excel_helper'

class ToExcelTest < Test::Unit::TestCase
  
  def setup
    @users = [
      User.new(:id => 1, :name => 'Dupont', :age => 25),
      User.new(:id => 2, :name => 'Martin', :age => 22)
    ]
    @computers =[
      Computer.new(:id => 1, :brand => 'Apple', :portable => true, :user => @users.first),
      Computer.new(:id => 2, :brand => 'Dell', :portable => false, :user => @users[1])
    ]
  end

  def test_should_be_array_of_arrays_not_empty
    assert_equal build_document(nil), [].to_excel
    assert_equal build_document(nil), [[]].to_excel
    assert_equal build_document(nil), @users.to_excel
  end

  def test_with_no_headers
    document = build_document(
      row(cell('Number', '25'), cell('Number', '1'), cell('String', 'Dupont')),
      row(cell('Number', '22'), cell('Number', '2'), cell('String', 'Martin'))
    )
    assert_equal document, [@users].to_excel(:headers => false, :map => {User => [:age, :id, :name]})
  end

  def test_with_methods
    document = build_document(
      row(cell('String', 'Age'), cell('String', 'Id'), cell('String', 'Name'), cell('String', 'Is old?')),
      row(cell('Number', '25'),  cell('Number', '1'),  cell('String', 'Dupont'),  cell('String', 'false')),
      row(cell('Number', '22'),  cell('Number', '2'),  cell('String', 'Martin'), cell('String', 'false'))
    )
    assert_equal document, [@users].to_excel(:map => { User => [:age, :id, :name, :is_old?]})
  end
  
  def test_with_nested_attribute
    document = build_document(
      row(cell('String', 'Id'), cell('String', 'Username')),
      row(cell('Number', '1'),  cell('String', 'Dupont')),
      row(cell('Number', '2'),  cell('String', 'Martin'))
    )
    assert_equal document, [@computers].to_excel(:map => { Computer => [:id, [:user, :name]]})
  end
  
  def test_with_two_collections
    document = build_document(
      row(cell('String', 'Id'), cell('String', 'Name'), cell('String', 'Id'), cell('String', 'Brand')),
      row(cell('Number', '1'),  cell('String', 'Dupont'),  cell('Number', '1'),  cell('String', 'Apple')),
      row(cell('Number', '2'),  cell('String', 'Martin'),  cell('Number', '2'),  cell('String', 'Dell'))
    )
    assert_equal document, [@users, @computers].to_excel(:map => { User => [:id,:name], Computer => [:id,:brand]})
  end
  
  def test_with_one_collection
    document = build_document(
      row(cell('String', 'Id'), cell('String', 'Username')),
      row(cell('Number', '1'),  cell('String', 'Dupont')),
      row(cell('Number', '2'),  cell('String', 'Martin'))
    )
    assert_equal document, [@computers].to_excel(:map => {Computer => [:id,[:user, :name]]})
  end
  
  def test_should_render_empty_cells_unless_collections_of_same_size
    @users << User.new(:id => 3, :name => 'Durand', :age => 23)
    document = build_document(
      row(cell('String', 'Id'), cell('String', 'Name'), cell('String', 'Id'), cell('String', 'Brand')),
      row(cell('Number', '1'),  cell('String', 'Dupont'),  cell('Number', '1'),  cell('String', 'Apple')),
      row(cell('Number', '2'),  cell('String', 'Martin'),  cell('Number', '2'),  cell('String', 'Dell')),
      row(cell('Number', '3'),  cell('String', 'Durand'),  cell('String', ''),  cell('String', ''))
    )
    assert_equal document, [@users, @computers].to_excel(:map => { User => [:id,:name], Computer => [:id,:brand]})
  end

  def test_with_own_header
    document = build_document(
      row(cell('String', 'User Id'), cell('String', 'Nom'), cell('String', 'Computer Id'), cell('String', 'Marque')),
      row(cell('Number', '1'),  cell('String', 'Dupont'),  cell('Number', '1'),  cell('String', 'Apple')),
      row(cell('Number', '2'),  cell('String', 'Martin'),  cell('Number', '2'),  cell('String', 'Dell'))
    )
    header = ['User Id', 'Nom', 'Computer Id', 'Marque']
    assert_equal document, [@users, @computers].to_excel(:map => { User => [:id,:name], Computer => [:id,:brand]}, :headers => header)    
  end
  #TODO generate default map from class columns if no map option passed
  def test_with_no_map_option

  end
  
private

  def build_document(*rows)
    "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Workbook xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\" xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:o=\"urn:schemas-microsoft-com:office:office\"><Worksheet ss:Name=\"Sheet1\"><Table>#{rows}</Table></Worksheet></Workbook>"
  end

  def row(*cells)
    "<Row>#{cells}</Row>"
  end

  def cell(type, content)
    "<Cell><Data ss:Type=\"#{type}\">#{content}</Data></Cell>"
  end
end
