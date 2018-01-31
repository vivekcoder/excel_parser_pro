require 'zip/filesystem'
require 'nokogiri'

module ExcelParser
  class Error < StandardError; end
  OOXMLEpoch = Date.new(1899,12,30)
  SharedStringPath = 'xl/sharedStrings.xml'
  StylesPath = 'xl/styles.xml'

  class Time < Struct.new(:hours, :minutes, :seconds)
  end

end

class ExcelParser::Workbook
  FORMATS = {
    'yyyy/mm/dd' => :date
  }

  STANDARD_FORMATS = {
    0 => 'General',
    1 => '0',
    2 => '0.00',
    3 => '#,##0',
    4 => '#,##0.00',
    9 => '0%',
    10 => '0.00%',
    11 => '0.00E+00',
    12 => '# ?/?',
    13 => '# ??/??',
    14 => 'mm-dd-yy',
    15 => 'd-mmm-yy',
    16 => 'd-mmm',
    17 => 'mmm-yy',
    18 => 'h:mm AM/PM',
    19 => 'h:mm:ss AM/PM',
    20 => 'h:mm',
    21 => 'h:mm:ss',
    22 => 'm/d/yy h:mm',
    37 => '#,##0 ;(#,##0)',
    38 => '#,##0 ;[Red](#,##0)',
    39 => '#,##0.00;(#,##0.00)',
    40 => '#,##0.00;[Red](#,##0.00)',
    45 => 'mm:ss',
    46 => '[h]:mm:ss',
    47 => 'mmss.0',
    48 => '##0.0E+0',
    49 => '@',
  }

  def initialize(file, user_defined_formats = {})
    @file = file
    zipfile = Tempfile.new("file")
    zipfile.binmode
    zipfile.write(HTTParty.get(@file).body)
    zipfile.close
    @file = zipfile.path
    @zipfs = Zip::File.open(@file)
    @user_defined_formats = user_defined_formats
    read_styles
  end

  def sheets
    begin
      workbook = Nokogiri::XML::Document.parse(@zipfs.file.open('xl/workbook.xml'))
    rescue Zip::Error
      raise ExcelParser::Error, 'Invalid file, could not open xl/workbook.xml'
    end
    @sheets = workbook.css('sheet').each_with_index.map do |n, i|
      ExcelParser::Sheet.new(self, n.attr('name'), n.attr('sheetId'), i+1)
    end
  end

  def string_table
    @string_table ||= read_string_table
  end

  def read_string_table
    return [] unless @zipfs.file.exist? ExcelParser::SharedStringPath

    begin
      shared_string = @zipfs.file.open(ExcelParser::SharedStringPath)
    rescue Zip::Error
      raise ExcelParser::Error, 'Invalid file, could not open shared string file.'
    end

    entry = ''
    @string_table = []
    Nokogiri::XML::Reader(shared_string).each do |node|
      if node.name == "si" and node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
        entry = ''
      elsif node.name == "si" and node.node_type == Nokogiri::XML::Reader::TYPE_END_ELEMENT
        @string_table << entry
      elsif node.value?
        entry << node.value
      end
    end
    @string_table
  end

  def read_styles
    @num_formats = {}
    @cell_xfs = []
    return unless @zipfs.file.exist? ExcelParser::StylesPath

    begin
      doc = Nokogiri::XML(@zipfs.file.open(ExcelParser::StylesPath))
    rescue Zip::Error
      raise ExcelParser::Error, 'Invalid file, could not open styles'
    end

    doc.css('/styleSheet/numFmts/numFmt').each do |numFmt|
      if numFmt.attributes['numFmtId'] && numFmt.attributes['formatCode']
        numFmtId = numFmt.attributes['numFmtId'].value.to_i
        formatCode = numFmt.attributes['formatCode'].value
        @num_formats[numFmtId] = formatCode
      end
    end

    doc.css('/styleSheet/cellXfs/xf').each do |xf|
      if xf.attributes['numFmtId']
        numFmtId = xf.attributes['numFmtId'].value.to_i
        @cell_xfs << numFmtId
      end
    end
  end

  def attribute_to_type(t, s)
    if t == 's'
      :shared
    elsif t == 'b'
      :boolean
    else
      id = @cell_xfs[s.to_i].to_i
      result = @num_formats[id]

      if result == nil
        if STANDARD_FORMATS.has_key? id
          result = STANDARD_FORMATS[id]
        end
      end
      format = result.downcase.sub('\\', '')

      if @user_defined_formats.has_key? format
        @user_defined_formats[format]
      else
        FORMATS[format] || :string
      end
    end
  end

  def zipfs
    @zipfs
  end

  def close
    @zipfs.close
  end
end

class ExcelParser::Sheet
  attr_reader :name, :workbook
  def initialize(workbook, name, id, index)
    @workbook = workbook
    @name = name
    @id = id
    @index = index
    begin
      @file = @workbook.zipfs.file.open(path) if @workbook.zipfs.file.exist?(path)
    rescue Zip::Error => e
      raise ExcelParser::Error, "Couldn't open sheet #{index}: #{e.message}"
    end
  end

  def string_lookup(i)
    @workbook.string_table[i] || (raise ExcelParser::Error, 'File invalid, invalid string table.')
  end

  def rows
    Enumerator.new(row_count) do |y|
      next unless @file
      @file.rewind
      shared = false
      row = nil
      cell_map = nil # Map of column letter to cell value for a row
      column = nil
      cell_type = nil
      Nokogiri::XML::Reader(@file).each do |node|
        case node.node_type
        when Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.name
          when "row"
            cell_map = {}
            next
          when 'c'
            node_type = node.attributes['t']
            node_style = node.attributes['s']
            cell_index = node.attributes['r']
            if !cell_index
              raise ExcelParser::Error, 'Invalid spreadsheet XML.'
            end
            column = cell_index.delete('0-9')
            cell_type = @workbook.attribute_to_type(node_type, node_style)
            shared = (node_type == 's')
            next
          end
        when Nokogiri::XML::Reader::TYPE_END_ELEMENT
          if node.name == 'row'
            y << process_row(cell_map)
          end
          next
        end

        if node.value
          value = case cell_type
            when :shared
              string_lookup(node.value.to_i)
            when :boolean
              node.value.to_i != 0
            when :datetime, :date
              ExcelParser::OOXMLEpoch + node.value.to_f
            when :time
              parse_time(node.value.to_f)
            when :float
              node.value.to_f
            else
              # leave as string
              node.value
            end
          cell_map[column] = value
        end
      end
    end
  end

  def parse_time(float)
    hours = (float * 24).floor
    minutes = (float * 24 * 60).floor % 60
    seconds = (float * 24 * 60 * 60).floor % 60
    ExcelParser::Time.new(hours, minutes, seconds)
  end

  def process_row(cell_map)
    max = cell_map.keys.map {|c| self.class.column_name_to_index c }.max
    row = []
    self.class.column_names[0..max].each do |col|
      if self.class.column_name_to_index(col) > max
        break
      else
        row << cell_map[col]
      end
    end
    row
  end



  # Returns A to ZZZ.
  def self.column_names
    if @column_names
      @column_names
    else
      proc = Proc.new do |prev|
        ("#{prev}A".."#{prev}Z").to_a
      end
      x = proc.call("")
      y = x.map(&proc).flatten
      z = y.map(&proc).flatten
      @column_names = x + y + z
    end
  end

  def self.column_name_to_index(name)
    if not @column_names_to_indices
      @column_names_to_indices = {}
      self.column_names.each_with_index do |name, i|
        @column_names_to_indices[name] = i
      end
    end
    @column_names_to_indices[name]
  end

  def row_count
    if defined? @row_count
      @row_count
    elsif @file
      @file.rewind
      Nokogiri::XML::Reader(@file).each do |node|
        if node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.name
          when 'dimension'
            if ref = node.attributes["ref"]
              break @row_count = ref.scan(/\d+$/).first.to_i
            end
          when 'sheetData'
            break @row_count = nil
          end
        end
      end
    end
  end

  private
  def path
    "xl/worksheets/sheet#{@index}.xml"
  end

end