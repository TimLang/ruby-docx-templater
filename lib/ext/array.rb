
class Object
  def deep_map!(&block); yield(self) end
end

class Array
  def deep_map!(&block); map!{|ary| ary.deep_map!(&block)} end
end
