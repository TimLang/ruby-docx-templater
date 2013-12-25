
require 'debugger'

class Object
  def deep_map!(&block); yield(self) end
end

class Array
  def deep_map!(&block); map!{|ary| ary.deep_map!(&block)} end

  #the data structure for exporting word
  def to_word_table
    return if self.empty? || !self[0].is_a?(Array)
    row_regex = /_\${(\d+)}\$_/
    col_regex = /__{(\d+)}__/
    cache = {}

    self.each_with_index do |rows, row_index|
      rows.each_with_index do |col, col_index|
        if col =~ row_regex
          ($1.to_i).times do |index|
            next if index < 1
            key = "#{row_index+index}"
            inc = rows[0..col_index].select{|r| r=~col_regex}.map{|a| col_regex.match(a)[1].to_i}.inject(&:+).to_i
            final_index = (inc==0 ? 0 : inc-1) + col_index
            if cache[key]
              cache[key] << final_index
            else
              cache[key] = [final_index]
            end
          end
        end
      end
    end

    cache.each do |k, v|
      v.each{|obj| self[k.to_i].insert(obj.to_i, '{_}')}
    end if cache

    self.map(&:compact)
  end

end
