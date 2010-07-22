require 'win32/registry'

class WIN32OLE_TYPELIB
  module TypeLibHelper
    private

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
          puts "GUID #{guid} #{version} #{arch} #{type_lib}"
          return type_lib if type_lib && name == typelib_name
        end
      end
      nil
    end
  end

  attr_reader :typelib
  attr_reader :name
  alias :to_s :name

  def initialize(*args)
    # TODO: Make this work internally and externally API w/ regards to inargs
    if args.length == 2
      @typelib, @name = *args
      puts "NO TYPELIB! for #{@name} #{@version}" unless @typelib
    elsif args.length == 1
      @name = args[0]
      @typelib = search_registry(@name) # TODO: Missing search_registry2
      puts "NAME IS #{@name}///#{@typelib}"
    end
  end

  def guid
    @typelib.guid
  end

  def minor_version
    @typelib.minor_version
  end

  def major_version
    @typelib.major_version
  end

  def ole_classes # MRI: ole_types_from_typelib
    ole_classes = []
    @typelib.type_info_count.times do |i|
      docs = @typelib.get_documentation(i)
      next unless docs
      info = @typelib.get_type_info(i)
      next unless info
      ole_classes << WIN32OLE_TYPE.new(self, info, docs)
    end
    ole_classes
  end

  def version
    [minor_version, major_version].join('.')
  end

  def inspect
    name
  end

  class << self
    def ole_classes(typelib)
      new(typelib).ole_classes
    end

    def typelibs
      typelibs = []
      typelib_registry_each_guid_version do |guid, version, reg|
        name = reg.read(nil)[1] || ''
        registry_subkey(reg, 'win32', 'win64') do |arch_reg, arch|
          type_lib = load_typelib(arch_reg, arch)
          # TODO: I think MRI figures out a few more typelibs than we do
          typelibs << WIN32OLE_TYPELIB.new(type_lib, name) if type_lib
        end
      end
      typelibs
    end

    include TypeLibHelper
  end

  include TypeLibHelper
end
