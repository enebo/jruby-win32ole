class WIN32OLE
  module Utils
    def SafeStringValue(str)
      return str if str.kind_of?(::String)
      if str.respond_to?(:to_str)
        str = str.to_str
        return str if str.kind_of?(::String)
      end
      raise TypeError
    end
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

    def load_typelib(path_reg, arch)
      path = path_reg.open(arch) { |r| r.read(nil) }[1]
#      puts "PATH = #{path}"
      begin
        Automation.loadTypeLib(path)
      rescue ComFailException => e
        puts "Failed to load #{name} fom #{path} because: #{e}"
        nil
      end
    end

    def find_all_typeinfo(typelib)
      typelib.type_info_count.times do |i|
        docs = typelib.get_documentation(i)
        next unless docs
        info = typelib.get_type_info(i)
        next unless info
        yield info, docs
      end      
    end

    def reg_each_key_for(reg, subkey, &block)
      reg.open(subkey) do |subkey_reg|
        subkey_reg.each_key { |key, wtime| block.call(subkey_reg, key) }
      end
    end

    # Walks all guid/clsid entries and yields every single version
    # of those entries to the supplied block. See search_registry as 
    # an example of its usage.
    def typelib_registry_each_guid_version
      Win32::Registry::HKEY_CLASSES_ROOT.open('TypeLib') do |reg| 
        reg.each_key do |guid, wtime|
          reg.open(guid) do |guid_reg|
            guid_reg.each_key do |version_string, wtime|
              version = version_string.to_f
              begin
                guid_reg.open(version_string) do |version_reg|
                  yield guid, version, version_reg
                end
              rescue Win32::Registry::Error => e
                # Version Entry may not contain anything. Skip.
              end
            end
          end
        end
      end
    end

    def registry_subkey(reg, *valid_subkeys)
      reg.each_key do |inner, wtime|
        reg_each_key_for(reg, inner) do |subkey_reg, subkey|
          yield subkey_reg, subkey if valid_subkeys.include? subkey
        end
      end
    end

    def search_registry(typelib_name) # MRI: oletypelib_search_registry
      typelib_registry_each_guid_version do |guid, version, reg|
        name = reg.read(nil)[1] || ''
        registry_subkey(reg, 'win32', 'win64') do |arch_reg, arch|
          type_lib = load_typelib(arch_reg, arch)
#          puts "GUID #{guid} #{version} #{arch} #{type_lib}"
          return type_lib if type_lib && name == typelib_name
        end
      end
      nil
    end
  end
end
