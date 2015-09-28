#********************************************************************************************
# Copyright (c) 2013 - 2014, Freescale Semiconductor, Inc.
# All rights reserved.
# #
# Redistribution and use in source and binary forms, with or without modification,
# are permitted provided that the following conditions are met:
# #
# o Redistributions of source code must retain the above copyright notice, this list
#   of conditions and the following disclaimer.
# #
# o Redistributions in binary form must reproduce the above copyright notice, this
#   list of conditions and the following disclaimer in the documentation and/or
#   other materials provided with the distribution.
# #
# o Neither the name of Freescale Semiconductor, Inc. nor the names of its
#   contributors may be used to endorse or promote products derived from this
#   software without specific prior written permission.
# #
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
# ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR
# ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
# ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
# SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
# Description: Use to scan the pojects of IAR, KEIL, KDS, GCC and Atollic
# in the KSDK package, generate the matrix and total count.
# Note that, it suppors the KSDK 1.2.0 GA or higher.
#
# Revision History:
# -----------------
# Code Version    YYYY-MM-DD    Author        Description
# 0.1             2015-02-02    Bill Yuan     Create this file
# 0.2             2015-02-10    Bill Yuan     Save to an excel by default
# 0.3             2015-02-10    Bill Yuan     Integration lib scanning
# 0.4             2015-02-15    Bill Yuan     Export to excel automatically
#********************************************************************************************
require 'optparse'
require 'find'
require 'win32ole'
require 'pathname'

DEBUG = false

#PLATFORM_LIST = ["frdmk82f", "twrk80f150m", "twrk81f150m"]
PLATFORM_LIST = ["All"]

#TOOL_CHAIN_LIST = ["ATL"]
TOOL_CHAIN_LIST = ["IAR", "KEIL", "GCC", "KDS", "ATL"]

CONFIGS_OPTION = ["Debug", "Release"]

INCLUDE_LINUX_VERSION         = true

INCLUDE_KW01_EXTERNAL_VERSION = true

class ProjectMatrix
  attr_reader :type, :root_path, :case_suit

  def process_cmdargs()
    opt_parser = OptionParser.new do | opts |
      opts.on("-t", "--type [case type]", String, \
        "case type: 'demo' or 'example' or 'usb' \n\
         \t\t\t\tor 'mqx' or 'mqx_mfs' or 'mqx_rtcs' or 'lib' \n") do | value |
          if value == nil
            puts(opts)
            exit(0)
          end
          @type = value.downcase
      end
      opts.on("-d", "--dir [root directory path]", String, \
        "the root directory path of KSDK\n") do | value |
          if value == nil
            puts(opts)
            exit(0)
          end
          root = value.gsub(/\\/,'/')
          if (!Dir.exist?(root))
            puts "Root directory does not exist, plesae double check: #{value}"
            exit(0)
          end
          # remove last duplicated / in the path
          @root_path = root.sub(/\/$/,'')
      end
      # help option - print help and ends
      opts.on("-h", "--help", "print this help\n\n") do
        puts(opts)
        puts "Example: ruby scan_projects.rb <type> <dir>\n"
        puts ""
        puts "\t ruby scan_projects.rb demo C:\\KSDK_1.2.0-GA_RC1\\KSDK_1.2.0\n"
        exit(0)
      end
    end

    opt_parser.parse!

    if DEBUG
      puts "type: #{@type}"
      puts "root_path: #{@root_path}"
    end

    # no entry file - print help and ends
    if (@type == nil || (@type != 'demo' && @type != 'example' &&  @type != 'usb'\
        &&  @type != 'mqx' &&  @type != 'mqx_mfs' && @type != 'mqx_rtcs' && @type != 'lib')\
     || @root_path == nil)
      puts(opt_parser)
      exit(0)
    end
  end

  def collect_info()
    waiting_thread = Thread.new do
      puts "Scanning the projects"
      100.times do
        sleep 1
        print(".")
      end
    end

    @case_suit = eval(@type.capitalize + 'Suit').new(@type, @root_path)
    @case_suit.scan()

    waiting_thread.exit if waiting_thread != nil

    @case_suit.output_info()
  end

  def generate_report()
    report_engine = ReportEngine.new(@type, @case_suit.get_projects, @case_suit.get_platforms,\
       @case_suit.get_toolchains, @case_suit.get_configs, @case_suit.get_case_matrix)

    report_engine.create_table()
    report_engine.fill_project_column()
    report_engine.fill_platforms_toolchains_column()
    report_engine.fill_matrix()
    report_engine.close()
  end

  class CaseSuit
    attr_accessor :case_type, :root_path, :folder_path, :projects, :platforms, :toolchains,
                  :filtter_folder, :no_filtter_folder, :case_matrix

      PROJECTS_ROOT_FOLDER_NAME = "examples"

    def initialize(type, path)
      @case_type = type
      @root_path = path
      @folder_path = "#{@root_path}/#{PROJECTS_ROOT_FOLDER_NAME}" #all projects are in this folder except MQX
      @projects = Array.new
      @platforms = Array.new
      @toolchains = Array.new
      @case_matrix = Hash.new
      @no_filtter_folder = ""
    end

    def get_projects()
      return @projects
    end

    def get_platforms()
      return @platforms
    end

    def get_toolchains()
      return @toolchains
    end

    def get_configs()
      return CONFIGS_OPTION
    end

    def get_case_matrix()
      return @case_matrix
    end

    def scan()
      puts "Start scan()" if DEBUG
      puts "folder_path: #{folder_path}" if DEBUG

      project_name = nil
      platform_name = nil
      toolchain_name = nil
      configs = Array.new()

      toolchains_filtter = Array.new()
      TOOL_CHAIN_LIST.each do |tool_chain|
        toolchains_filtter.push(eval(tool_chain.upcase() + 'ToolChainParser').new())
      end

      # scan all files, include the hide file(.cporject)
      Dir.glob("#{@folder_path}/**/{.*,*}") do |path|

        if @case_type != 'lib'
          in_platform_list = false
          PLATFORM_LIST.each do |platform|
            if (platform != nil && platform.upcase == "ALL") || path =~ /\/#{platform}\//
              in_platform_list = true
              break
            end
          end
          next if !in_platform_list
        end

        # only scan the sepecial
        if path =~ /\/#{@filtter_folder}\//

          # skip the no filtter folder, becuase usb has put into examples folder
          if path =~ /\/#{@no_filtter_folder}\//
            next
          end

          toolchains_filtter.each do |toolchain|
            if !File.directory?("#{path}") && toolchain.is_toolchain(path)
              puts "toolchain parser path is: #{path}" if DEBUG
              platform_name = get_platform_name(path)
              project_name = toolchain.get_project_name(path, platform_name)
              toolchain_name = toolchain.get_name()
              configs = Array.new()
              CONFIGS_OPTION.each do |config|
                configs.push(config) if toolchain.has_config(config, path, @case_type)
              end
              puts "Config list is: #{configs}" if DEBUG
              # skip the other tool chains scanning
              break
            end
          end

          if project_name != nil && platform_name != nil && toolchain_name != nil && !configs.empty?
            # add to hash
            add_project(project_name)

            add_platform(project_name, platform_name)
            if INCLUDE_KW01_EXTERNAL_VERSION && platform_name == "mrbkw01"
              add_platform(project_name, "mrbkw01_ja")
              add_platform(project_name, "mrbkw01_eu")
            end

            add_toolchain(toolchain_name)
            if INCLUDE_LINUX_VERSION && KDSToolChainParser::TOOL_CHAIN_NAME == toolchain_name
              add_toolchain(KDSToolChainParser::TOOL_CHAIN_NAME_LINUX)
            elsif INCLUDE_LINUX_VERSION && GCCToolChainParser::TOOL_CHAIN_NAME == toolchain_name
              add_toolchain(GCCToolChainParser::TOOL_CHAIN_NAME_LINUX)
            end

            set_matrix(project_name, platform_name, toolchain_name, configs)
            if INCLUDE_KW01_EXTERNAL_VERSION && platform_name == "mrbkw01"
              set_matrix(project_name, "mrbkw01_ja", toolchain_name, configs)
              set_matrix(project_name, "mrbkw01_eu", toolchain_name, configs)
            end
          end
        end
      end
    end

    def add_project(project_name)
      if !@case_matrix.has_key?(project_name)
        @case_matrix[project_name] = Hash.new()
        @projects.push(project_name) unless @projects.include?(project_name)
      end
    end

    def add_platform(project_name, platform_name)
      if !@case_matrix[project_name].has_key?(platform_name)
        @case_matrix[project_name][platform_name] = Hash.new()
        @platforms.push(platform_name) unless @platforms.include?(platform_name)
      end
    end

    def add_toolchain(toolchain_name)
      #adjust tool chains priority
      if !@toolchains.include?(toolchain_name)
        if IARToolChainParser::TOOL_CHAIN_NAME == toolchain_name
          @toolchains.insert(0, toolchain_name) unless @toolchains.include?(toolchain_name)
        elsif ATLToolChainParser::TOOL_CHAIN_NAME == toolchain_name
          @toolchains.insert(((@toolchains.length - 1) >= 0 ? -1 : 0), toolchain_name) unless @toolchains.include?(toolchain_name)
        else
          # put to the beginning for the KDS
          @toolchains.insert(((@toolchains.length - 2) >= 0 ? -2 : 0), toolchain_name) unless @toolchains.include?(toolchain_name)
        end
      end
    end

    def set_matrix(project_name, platform_name, toolchain_name, configs)
      if !@case_matrix[project_name][platform_name].has_key?(toolchain_name)
        @case_matrix[project_name][platform_name][toolchain_name] = Hash.new()
      end
      if INCLUDE_LINUX_VERSION && KDSToolChainParser::TOOL_CHAIN_NAME == toolchain_name
        if !@case_matrix[project_name][platform_name].has_key?(KDSToolChainParser::TOOL_CHAIN_NAME_LINUX)
          @case_matrix[project_name][platform_name][KDSToolChainParser::TOOL_CHAIN_NAME_LINUX] = Hash.new()
        end
      end

      if INCLUDE_LINUX_VERSION && GCCToolChainParser::TOOL_CHAIN_NAME == toolchain_name
        if !@case_matrix[project_name][platform_name].has_key?(GCCToolChainParser::TOOL_CHAIN_NAME_LINUX)
          @case_matrix[project_name][platform_name][GCCToolChainParser::TOOL_CHAIN_NAME_LINUX] = Hash.new()
        end
      end

      configs.each do |config|
        if !@case_matrix[project_name][platform_name][toolchain_name].has_key?(config)
          @case_matrix[project_name][platform_name][toolchain_name][config] = true

          #add linux version, duplicated windows version
          if INCLUDE_LINUX_VERSION && KDSToolChainParser::TOOL_CHAIN_NAME == toolchain_name
            @case_matrix[project_name][platform_name][KDSToolChainParser::TOOL_CHAIN_NAME_LINUX][config] = true
          end

          if INCLUDE_LINUX_VERSION && GCCToolChainParser::TOOL_CHAIN_NAME == toolchain_name
            @case_matrix[project_name][platform_name][GCCToolChainParser::TOOL_CHAIN_NAME_LINUX][config] = true
          end
        end
      end
    end

    def get_platform_name(path)
      # the default platform name is followed by /apps/, such as \mcu-sdk\apps\'twrk21d50m'\examples
      exclude_root = path.sub(/^#{@root_path}\/#{PROJECTS_ROOT_FOLDER_NAME}\//,'')
      puts "exclude_root: #{exclude_root}" if DEBUG

      platform_name = exclude_root.split("/").first
      puts "get_platform_name: #{platform_name}" if DEBUG
      return platform_name
    end

    def output_info
      puts "\n------------projects:#{@projects.count}-------------\n"
      puts @projects
      puts "\n------------platforms:#{@platforms.count}------------\n"
      puts @platforms
      puts "\n------------toolchains:#{@toolchains.count}------------\n"
      puts @toolchains
      puts "\n------------configs:#{CONFIGS_OPTION.count}------------\n"
      puts CONFIGS_OPTION
      puts
    end
  end

  class DemoSuit < CaseSuit
    def initialize(type, path)
      super(type, path)
      @filtter_folder = "demo_apps"
      @no_filtter_folder = "usb"
    end
  end

  class LibSuit < CaseSuit
    def initialize(type, path)
        super(type, path)
        @root_path = "#{path}"
        @folder_path = "#{@root_path}/lib"
        @filtter_folder = "lib"
    end
    def get_platform_name(path)
      if !File.directory?(path) && !path.match(/debug/) && !path.match(/release/)
          platform_name = File.dirname(path).split("/").last.split("_").last
          puts "get_platform_name: #{platform_name}, #{path}" if DEBUG
          return platform_name
      else
          return nil
      end
    end
  end

  class ExampleSuit < CaseSuit
    def initialize(type, path)
      super(type, path)
      @filtter_folder = "driver_examples"
    end
  end

  class UsbSuit < CaseSuit
    def initialize(type, path)
      super(type, path)
      @filtter_folder = "usb"
    end
  end

  class MqxSuit < CaseSuit
    def initialize(type, path)
      super(type, path)
      @root_path = "#{path}/rtos/mqx/mqx"
      @folder_path = "#{@root_path}/examples" #all projects are in this folder except MQX
      @filtter_folder = "examples"
    end

    def get_platform_name(path)
      platform_name = File.dirname(path).split("/").last.split("_").last

      # "special case for usbkw40, need append the last - 1 to the name"
      if "k22f" == platform_name or "kw40z" == platform_name
        platform_name = File.dirname(path).split("/").last.split("_")[-2] + "_" + platform_name
      end

      puts "get_platform_name: #{platform_name}" if DEBUG
      return platform_name
    end
  end

  class Mqx_mfsSuit < MqxSuit
    def initialize(type, path)
      super(type, path)
      @root_path = "#{path}/middleware/filesystem/mfs"
      @folder_path = "#{@root_path}/examples"
    end
  end

  class Mqx_rtcsSuit < MqxSuit
    def initialize(type, path)
      super(type, path)
      @root_path = "#{path}/middleware/tcpip/rtcs"
      @folder_path = "#{@root_path}/examples"
    end
  end

  class ToolChainParser
    attr_accessor :name, :tool_chain_filtter

    def is_toolchain(path)
      # need include the toolchain name in its path, such as XXX/iar/XXX/XXX.ewp
      if path =~ /\/#{@name}\//
        return (path =~ /#{@tool_chain_filtter}$/) != nil
      else
        return false
      end
    end

    def get_name()
      return @name
    end

    def get_project_name(path, platform_name)
      # special for USB project, need exclude project name, the format of usb project is XXX_[platform_name]
      path = path.gsub(/_#{platform_name}/, '')

      # default way is get the project name according to its file name, such as XXXX/adc_pit_trigger.eww
      project_name = path.split("/").last.split(".#{@tool_chain_filtter}").first
      puts "get_project_name: #{project_name}, path is #{path}" if DEBUG
      return project_name
    end

    def has_config(config, file, type)
      return false
    end

    def self.filtter_config(filtter_str, file)
      if !File.exist?(file)
        puts "not found the config file: #{file}"
        exit(1)
      end

      puts "filtter str: #{filtter_str}" if DEBUG
      config_file = File.open(file, :encoding => 'utf-8')
      line = config_file.grep(/#{filtter_str}/)
      config_file.close if config_file != nil

      if line != nil && line.length > 0
        return true
      end

      return false
    end
  end

  class IARToolChainParser < ToolChainParser
    TOOL_CHAIN_NAME = "iar"
    TOOL_CHAIN_FILTER = "ewp"

    def initialize()
      @name = TOOL_CHAIN_NAME
      @tool_chain_filtter = TOOL_CHAIN_FILTER
    end

    def has_config(config, file, type)
      filtter_str = "#{config}"
      if type != nil && type.downcase.start_with?("mqx")
        filtter_str = "\<name\>int flash #{config.downcase}\<\/name\>"
      else
        filtter_str = "\<name\>#{config}\<\/name\>"
      end

      ToolChainParser.filtter_config(filtter_str, file)
    end
  end

  class KEILToolChainParser < ToolChainParser
    TOOL_CHAIN_NAME = "mdk"
    TOOL_CHAIN_FILTER = "uvprojx"

    def initialize()
      @name = TOOL_CHAIN_NAME
      @tool_chain_filtter = TOOL_CHAIN_FILTER
    end

    def has_config(config, file, type)
      filtter_str = "#{config}"
      if type != nil && type.downcase.start_with?("mqx")
        filtter_str = "\<OutputDirectory\>int flash #{config.downcase}"
      else
        filtter_str = "\<OutputDirectory\>#{config.downcase}"
      end
      ToolChainParser.filtter_config(filtter_str, file)
    end
  end

  class KDSToolChainParser < ToolChainParser
    TOOL_CHAIN_NAME = "kds"
    TOOL_CHAIN_NAME_LINUX = "kds_linux"
    TOOL_CHAIN_FILTER = "cproject"
    TOOL_CHAIN_FILTER_LAUNCH = "launch"

    def initialize()
      @name = TOOL_CHAIN_NAME
      @tool_chain_filtter = TOOL_CHAIN_FILTER
    end

    def get_project_name(path, platform_name)
      project_name = nil
      if path.match(/lib/)
        #it's for lib projects scanning
        path = path.gsub(/.*\/lib\//, '')
        project_name = path.split("/").first
      else
        Dir.foreach(File.dirname(path)) do |sub_path|
          # according to its launch name to get the project name, such as XXX/adc_pit_trigger_frdmk64f debug jlink.launch
          puts "get_project_name, loop path: #{sub_path}" if DEBUG
          if sub_path =~ /.#{TOOL_CHAIN_FILTER_LAUNCH}$/
            # special for USB project, need exclude project name
            sub_path = sub_path.gsub(/_#{platform_name}/, '')

            project_name = sub_path.split("/").last.split(" ").first
            puts "get_project_name: #{project_name}" if DEBUG
            break
          end
        end
      end
      return project_name
    end

    def has_config(config, file, type)
      filtter_str = "#{config}"
      if type != nil && type.downcase.start_with?("mqx")
        filtter_str = "name=\"int flash #{config.downcase}\""
      else
        filtter_str = "\<configuration configurationName=\"#{config.downcase}\"\>"
      end
      ToolChainParser.filtter_config(filtter_str, file)
    end
  end

  class ATLToolChainParser < KDSToolChainParser
    TOOL_CHAIN_NAME = "atl"

    def initialize()
      @name = TOOL_CHAIN_NAME
    end

    def has_config(config, file, type)
      filtter_str = "#{config}"
      if type != nil && type.downcase.start_with?("mqx")
        filtter_str = "\<configuration \[\\s\\S\]\* name=\"int flash #{config.downcase}\""
      else
        filtter_str = "\<configuration \[\\s\\S\]\* name=\"#{config.downcase}\""
      end
      ToolChainParser.filtter_config(filtter_str, file)
    end
  end

  class GCCToolChainParser < ToolChainParser
    TOOL_CHAIN_NAME = "armgcc"
    TOOL_CHAIN_NAME_LINUX = "armgcc_linux"
    TOOL_CHAIN_FILTER = "CMakeLists.txt"

    def initialize()
      @name = TOOL_CHAIN_NAME
      @tool_chain_filtter = TOOL_CHAIN_FILTER
    end

    def get_project_name(path, platform_name)
      # need paser the make file to get the project name
      project_name = nil
      if path.match(/lib/)
        #it's for lib projects scanning
        path = path.gsub(/.*\/lib\//, '')
        project_name = path.split("/").first
      else
        File.open(path).each_line do |line|
          if line =~ /.elf/
            # for MQX makefile, it has '('' before the elf, while for ksdk, it has '/'
            project_name = line.split(".elf").first.split(" ").last.split(/[\/(\"]/).last

            if project_name != nil
              break #skip loop featch
            end
          end
        end
      end
      return project_name
    end

    def has_config(config, file, type)
      filtter_str = "#{config}"
      if type != nil && type.downcase.start_with?("mqx")
        filtter_str = "CMAKE_BUILD_TYPE MATCHES \"int flash #{config.downcase}\""
      else
        filtter_str = "CMAKE_BUILD_TYPE MATCHES #{config}"
      end
      ToolChainParser.filtter_config(filtter_str, file)
    end
  end

  class ReportEngine
    attr_reader :type_name, :projects, :platforms, :toolchains, :configs, :matrix, :excel, :workbook, :worksheet,
                :row_pos, :col_pos, :waiting_thread
    attr_accessor :total

    EXCEL_BASE_NAME           = 'projects_summary.xlsx'

    EXCEL_CELL_RANGE_START    = 65 #'A'
    EXCEL_CELL_RANGE_NUM      = 26

    PROJECT_NAME_START_COLUMN = 'A'
    PROJECT_NAME_START_ROW    =  5

    PLATFORM_NAME_ROW         =  2
    TOOLCHIAN_NAME_ROW        =  3
    CONFIG_NAME_ROW           =  4

    HORIZONTAL_ALIGNMENT_CODE = -4108
    CLOUMN_WIDTH              = 3.0

    MARK_COVERED_CASE         = 'Y'
    MARK_NOT_COVERED_CASE     = 'NA'
    MARK_NOT_COVERED_PLATFROM = 'NA'

    SUMMARY_ROW_OFFSET        = 1

    def initialize(name, case_projects, case_platforms, case_toolchains, case_configs, case_matrix)
      @type_name = name
      @projects = case_projects
      @platforms = case_platforms
      @toolchains = case_toolchains
      @configs = case_configs
      @matrix = case_matrix
      @total = 0
      @row_pos = Hash.new()
      @col_pos = Hash.new()
    end

    def create_table()
      @waiting_thread = Thread.new do
        puts "Generating the report"
        100.times do
          sleep 1
          print(".")
        end
      end

      @excel = WIN32OLE.new('Excel.Application')
      @workbook = excel.Workbooks.Add()
      @worksheet = workbook.Worksheets(1)
      @worksheet.Select
    end

    def fill_project_column()
      projects.each do |project|
        row = projects.index(project) + PROJECT_NAME_START_ROW
        @row_pos[project] = row
        @worksheet.Range("#{PROJECT_NAME_START_COLUMN}#{row}").value = project
      end
      worksheet.Columns("#{PROJECT_NAME_START_COLUMN}").Autofit
    end

    def fill_platforms_toolchains_column()
      @platforms.each do |platform|
        offset_platform = platforms.index(platform) * toolchains.count * configs.count + 1 #start from 'B'
        cell_name = get_column_name(offset_platform)
        worksheet.Range("#{cell_name}#{PLATFORM_NAME_ROW}").value = platform

        if !@col_pos.has_key?(platform)
          @col_pos[platform] = Hash.new
        end

        toolchains.each do |tool|
          offset_toolchain = toolchains.index(tool) * configs.count + offset_platform
          toolchain_cell_name = get_column_name(offset_toolchain)
          worksheet.Range("#{toolchain_cell_name}#{TOOLCHIAN_NAME_ROW}").value = tool

          if !@col_pos[platform].has_key?(tool)
            @col_pos[platform][tool] = Hash.new
          end

          configs.each do |config|
            offset_config = configs.index(config) + offset_toolchain
            config_cell_name = get_column_name(offset_config)
            worksheet.Range("#{config_cell_name}#{CONFIG_NAME_ROW}").value = config
            @col_pos[platform][tool][config] = config_cell_name
          end

          toolchain_cell_name_end = get_column_name(offset_toolchain + configs.count - 1)
          worksheet.Range("#{toolchain_cell_name}#{TOOLCHIAN_NAME_ROW}:#{toolchain_cell_name_end}#{TOOLCHIAN_NAME_ROW}").Merge
        end

        cell_name_end = get_column_name(offset_platform + toolchains.count * configs.count - 1)
        worksheet.Range("#{cell_name}#{PLATFORM_NAME_ROW}:#{cell_name_end}#{PLATFORM_NAME_ROW}").Merge
      end
      worksheet.Rows(PLATFORM_NAME_ROW).HorizontalAlignment = HORIZONTAL_ALIGNMENT_CODE
      worksheet.Rows(TOOLCHIAN_NAME_ROW).HorizontalAlignment = HORIZONTAL_ALIGNMENT_CODE
      worksheet.Columns("B:ZZ").ColumnWidth = CLOUMN_WIDTH
    end

    def fill_matrix()
      matrix.each_key do |project|
        platforms.each do |platform|
          toolchains.each do |tool|
            configs.each do |config|
              col = @col_pos[platform][tool][config]
              row = @row_pos[project]
              if matrix[project].has_key?(platform)
                if matrix[project][platform].has_key?(tool)
                  if matrix[project][platform][tool].has_key?(config)
                    worksheet.Range("#{col}#{row}").value = MARK_COVERED_CASE
                    @total += 1 # Summary total count
                  else
                    worksheet.Range("#{col}#{row}").value = MARK_NOT_COVERED_CASE
                  end
                else
                  worksheet.Range("#{col}#{row}").value = MARK_NOT_COVERED_CASE
                end
              else
                worksheet.Range("#{col}#{row}").value = MARK_NOT_COVERED_PLATFROM
              end
            end
          end
        end
      end
      # fill the summary
      summary_row = SUMMARY_ROW_OFFSET
      worksheet.Range("#{PROJECT_NAME_START_COLUMN}#{summary_row}").value = "total projects number:#{@total}"
    end

    def close()
      @waiting_thread.exit if @waiting_thread != nil
      puts ""
      puts "-------------------finished--------------------"
      puts "Total count: #{@total}"

      # save the excel
      current_folder = Pathname.new(File.dirname(__FILE__)).realpath
      save_excel_path = File.expand_path("#{@type_name}_#{EXCEL_BASE_NAME}", "#{current_folder}")
      save_excel_path = save_excel_path.gsub('/','\\')
      puts ""
      puts "Saving to: #{save_excel_path}"
      begin
        workbook.SaveAs("#{save_excel_path}")
        puts "Saved successfully!\n\n"
      rescue
        puts "Failed to save, make sure the excel is not opened!"
      ensure
        excel.Quit() if excel != nil
      end
    end

    def get_column_name(offset)
      column_prefix_name = nil
      column_suffix_name = nil

      #only deal with B~ZZ range, can support 135 platforms
      prefix_index = offset/EXCEL_CELL_RANGE_NUM
      suffix_index = offset%EXCEL_CELL_RANGE_NUM

      if prefix_index >= 1
        column_prefix_name = (EXCEL_CELL_RANGE_START + prefix_index - 1).chr
      end

      column_suffix_name = (EXCEL_CELL_RANGE_START + suffix_index).chr

      puts "get_column_name, offset is #{offset}, cell name is #{column_prefix_name}#{column_suffix_name}" if DEBUG
      return "#{column_prefix_name}#{column_suffix_name}"
    end
  end
end

matrix = ProjectMatrix.new()
matrix.process_cmdargs()
matrix.collect_info()
matrix.generate_report()
