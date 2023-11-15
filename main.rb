require 'roo'
# require 'roo-xls'
require './sheet'
require './column'

file_name = './test.xlsx'
# file_name = './test_xls.xls'

sheet = Sheet.new(file_name, 0)

# pp sheet.array
# 1.
pp sheet.to_array

# 2.
pp sheet.row(1)[2]
pp sheet.row(1)

# 3.
sheet.each do |cell|
  pp cell
end

# 5. a
pp sheet["prva kolona"]
pp sheet["indeks"]
# 5. b
pp sheet["prva kolona"][1]
# 5. c
pp sheet['prva kolona'][1]
sheet['prva kolona'][1] = 12345
pp sheet['prva kolona'][1]

# 6.
pp sheet.prvaKolona
pp sheet.prvaKolona.sum
pp sheet.prvaKolona.avg
pp sheet.indeks.rn123
# ostavi mi samo brojeve u nizu
pp sheet.prvaKolona.select { |cell| cell.is_a?(Numeric) }
# na svaki element mi dodaj " asdf"
pp sheet.prvaKolona.map { |e| "#{e} asdf" }
# ovo ce da ostavi samo brojeve i da ih pomnozi na kraju
pp sheet.prvaKolona.select { |cell| cell.is_a?(Numeric) }.reduce(1, :*)

# 8 i 9
sheet2 = Sheet.new(file_name, 1)
pp sheet + sheet2
# pp sheet - sheet2
