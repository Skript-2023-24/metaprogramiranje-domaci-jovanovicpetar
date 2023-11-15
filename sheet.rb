class Sheet
  include Enumerable
  attr_accessor(:file, :sheet, :array)

  def initialize(file, index)
    # @file - roo spreadsheet klasa
    @file = Roo::Spreadsheet.open(file)
    # @sheet - roo worksheet klasa
    @sheet = @file.sheet(index)
    @index = index
    @array = []
    init_array
  end

  def init_array
    current_column = 0
    contains_total_row = false
    (sheet.first_column..sheet.last_column).each do |i|
      @array << Column.new(self)
      (sheet.first_row..sheet.last_row).each do |j|
        cell = sheet.cell(j, i)
        next if cell.nil? || cell == ''

        contains_total_row = true if cell.respond_to?(:include?) && cell.include?('total')

        @array[current_column].cells << cell
      end
      current_column += 1
    end
    remove_total if contains_total_row
  end

  def to_array
    result = []
    @array.each do |column|
      result << column.cells
    end
    result
  end

  def remove_total
    @array.each do |column|
      column.cells.pop
    end
  end

  def row(index)
    to_array.transpose[index]
  end

  def each
    to_array.transpose.each do |row|
      row.each do |x|
        yield x
      end
    end
  end

  def[](col)
    col = match col

    @array.each do |column|
      return column if col == match(column.name)
    end
  end

  def method_missing(name)
    self[name.to_s]
  end

  def respond_to_missing?
    true
  end

  def match(name)
    (name.gsub ' ', '').downcase
  end

  def +(other)
    return "Can't operate" unless can_operate?(other)

    @array.length.times do |i|
      @array[i] += other.array[i]
    end
    self
  end

  def -(other)
    return "Can't operate" unless can_operate?(other)

    @array.length.times do |i|
      @array[i] -= other.array[i]
    end
    self
  end

  def can_operate?(other)
    @array.length.times do |i|
      return false if @array[i].name != other.array[i].name
    end
    true
  end
end
