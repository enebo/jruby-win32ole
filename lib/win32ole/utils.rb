class WIN32OLE
  module Utils
    def from_variant(value)
      object = VariantUtilities.variant_to_object(value)
      object.kind_of?(Dispatch) ? WIN32OLE.new(object) : object
    end

    def all_methods(typeinfo, *args, &block) # MRI: olemethod_from_typeinfo
      # Find method in this type.
      ret = find_all_methods_in(nil, typeinfo, *args, &block)
      return ret if ret

      # Now check all other type impls
      typeinfo.impl_types_count.times do |i|
        begin
          href = typeinfo.get_ref_type_of_impl_type(i)
          ref_typeinfo = typeinfo.get_ref_type_info(href)
          ret = find_all_methods_in(typeinfo, ref_typeinfo, *args, &block)
          return ret if ret
        rescue ComFailException => e
          puts "Error getting impl types #{e}"
        end
      end
      nil
    end

    # MRI: ole_method_sub
    def find_all_methods_in(old_typeinfo, typeinfo, *args, &block)
      typeinfo.funcs_count.times do |i|
        begin
          desc = typeinfo.get_func_desc(i)
          docs = typeinfo.get_documentation(desc.memid)
          ret = yield typeinfo, old_typeinfo, desc, docs, *args
          return ret if ret
        rescue ComFailException => e
          puts "Error getting method info #{e}"
        end
      end
      nil
    end

    def typeinfo_from_ole # MRI: typeinfo_from_ole
      typeinfo = @dispatch.type_info
      docs = typeinfo.get_documentation(-1)
      type_lib = typeinfo.get_containing_type_lib
      type_lib.get_type_info_count.times do |i|
        begin
          ti = type_lib.get_type_info(i)
          tdocs = type_lib.get_documentation(i)
          return typelib.get_type_info(i) if tdocs.name == docs.name
        rescue ComFailException => e
          # We continue on failure. 
        end
      end
      type_info # Actually MRI seems like it could fail in weird case
    end
  end
end
