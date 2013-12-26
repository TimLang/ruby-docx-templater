
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
            c_inc = rows[0..col_index].reject{|r| r=~col_regex}.size
            final_index = ((col_index == 0 && c_inc==0) ? 0 : inc-1) + c_inc
            step_match = col_regex.match(col)
            step = (step_match ? step_match[1] : 1).to_i
            if cache[key]
              cache[key] << [final_index, step]
            else
              cache[key] = [[final_index, step]]
            end
          end
        end
      end
    end

    cache.each do |k, v|
      v.each do |obj| 
        ary = self[k.to_i]
        if ary
          obj[0] = ary.size if ary.size - obj[0].to_i < 0
          ary.insert(obj[0].to_i, "{_}__{#{obj[1]}}__")
        end
      end
    end if cache

    self.map(&:compact)
  end

end
