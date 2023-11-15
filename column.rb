class Column
  attr_accessor(:cells)

  def initialize(sheet)
    @sheet = sheet
    @cells = []
  end

  def name
    cells[0]
  end

  def[](index)
    cells[index]
  end

  def []=(index, value)
    cells[index] = value
  end

  def sum
    cells.filter { |cell| cell.is_a?(Numeric) }.reduce(:+)
  end

  def avg
    sum / cells.length.to_f
  end

  def method_missing(name)
    index = -1
    curr = 0
    cells.each do |cell|
      index = curr if cell == name.to_s
      curr += 1
    end
    return @sheet.row(index) unless index == -1

    "Cell couldn't be found"
  end

  def select(&)
    cells.select(&)
  end

  def map(&)
    cells.map(&)
  end

  def reduce(*arg)
    cells.reduce(*arg)
  end

  def respond_to_missing?
    true
  end

  def +(other)
    @cells + other.cells
  end

  def -(other)
    @cells - other.cells
  end
end
