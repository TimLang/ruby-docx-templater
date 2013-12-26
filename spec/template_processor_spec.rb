# -*- encoding : utf-8 -*-
require 'spec_helper'

module DocxTemplater
  module TestData

    MOCK_DATA = lambda {
      data = OpenStruct.new({
        :subject => '答题卡'
      })
      choice = "[A]\n[B]\n[C]\n[D]"

      ary = lambda { |i|
        15.times.inject([]) do |r, a|
          next r << {"c_#{i}".to_sym => choice} if i%2 == 0
          r << {"c_#{i}".to_sym => @count+=1}
        end
      }

      @count = 0
      12.times do |i|
        data.__send__("c_#{i+1}=", ary.call(i+1))
      end

      data.other = ([{:code => 'xxxx'}])

      puts data.marshal_dump
      return data.marshal_dump
    }


    DATA = {
      :subject => '高考模拟卷一',
      #:teacher => '土豪金',
      :teacher => '土豪金高端大气上档次哈哈',
           :questions => [
        {
          #:title => '',
          :title => '一、选择题（本题包括21小题，每小题给出的四个选项中，有的只有一个选项正确，有的有多个选项正确，全部选对的得6分，选对但不全的得3分，有选错的得0分）',
          :items => [
            {
              :content => "1. _``table:[[\"A\",\"哈哈\",\"测试图片哎\"],[\"V\",\"B\",\"${image2}\"]]``_人在恐惧、紧张时，在内脏神经的支配下，肾上腺髓质释放的肾上腺素增多，该激素可用于心脏，使心率加快。下列叙述错误的是（    ）",
              :choice => ["27. A.=   B.<   C.>   ", "28. A.=   B.<   C.>   ", "29. A.=   B.<   C.>   ", "30. A.=   B.<   C.>   "]
            },
            {
              :content => "1. Part I Writing\nFor this part,\nyou are allowed 30 minutes${image2} to write a short essay entitled Nothing Succeeds ${image3}${image4}Without a Strong Will by commenting on the humorous saying, \"Quitting smoking is ${image5}the easiest thing in the world. I've done it hundreds of times.\" You should write at least 120 words but no more than 180 ${image6}words.(本题30分)",
              :choice => ["A.She wants to take the most direct way.", "B.She may be late for the football game.", "C.She is worried about missing her flight.", "D.She is currently caught in a traffic jam."]
            },
            {
              :content => "_``table:[[\"Column one and two__{2}__\", \"Column three\"], [\"A large cell\", \"a smaller cell__{2}___${2}$_\"], [\"Anthoere\",'{_}__{1}__']]``_",
              :choice => ["A.She wants to take the most direct way.", "B.She may be late for the football game.", "C.She is worried about missing her flight.", "D.She is currently caught in a traffic jam."]
            },
            {
              :content => '2. 番茄幼苗在缺镁的培养液中培养一段时间后，与对照组相比，其叶片光合作用强度下降，原因是（    ）',
              :choice => ["A.She wants to take the most direct way.", "B.She may be late for the football game.", "C.She is worried about missing her flight.", "D.She is currently caught in a traffic jam."]
            },
            {
              :content => '2. 番茄幼苗在缺镁的培养液中培养一段时间后，与对照组相比，其叶片光合作用强度下降，原因是（    ）',
              :choice => ["A.She wants to take the most direct way.", "B.She may be late for the football game.", "C.She is worried about missing her flight.", "D.She is currently caught in a traffic jam."]
            },
            {
              :content => '2. 番茄幼苗在缺镁的培养液中培养一段时间后，与对照组相比，其叶片光合作用强度下降，原因是（    ）',
              :choice => ["A.She wants to take the most direct way.", "B.She may be late for the football game.", "C.She is worried about missing her flight.", "D.She is currently caught in a traffic jam."]
            },
            {
              :content => '2. 番茄幼苗在缺镁的培养液中培养一段时间后，与对照组相比，其叶片光合作用强度下降，原因是（    ）',
              :choice => ["A.She wants to take the most direct way.", "B.She may be late for the football game.", "C.She is worried about missing her flight.", "D.She is currently caught in a traffic jam."]
            },
            {
              :content => '2. 番茄幼苗在缺镁的培养液中培养一段时间后，与对照组相比，其叶片光合作用强度下降，原因是（    ）',
              :choice => ["A.She wants to take the most direct way.", "B.She may be late for the football game.", "C.She is worried about missing her flight.", "D.She is currently caught in a traffic jam."]
            },
            {
              :content => '2. 番茄幼苗在缺镁的培养液中培养一段时间后，与对照组相比，其叶片光合作用强度下降，原因是（    ）',
              :choice => ["A.She wants to take the most direct way.", "B.She may be late for the football game.", "C.She is worried about missing her flight.", "D.She is currently caught in a traffic jam."]
            },
            {
              :content => '2. 番茄幼苗在缺镁的培养液中培养一段时间后，与对照组相比，其叶片光合作用强度下降，原因是（    ）',
              :choice => ["A.She wants to take the most direct way.", "B.She may be late for the football game.", "C.She is worried about missing her flight.", "D.She is currently caught in a traffic jam."]
            },
            {
              :content => '2. 番茄幼苗在缺镁的培养液中培养一段时间后，与对照组相比，其叶片光合作用强度下降，原因是（    ）',
              :choice => ["A.She wants to take the most direct way.", "B.She may be late for the football game.", "C.She is worried about missing her flight.", "D.She is currently caught in a traffic jam."]
            },
            {
              :content => '2. 番茄幼苗在缺镁的培养液中培养一段时间后，与对照组相比，其叶片光合作用强度下降，原因是（    ）',
              :choice => ["A.She wants to take the most direct way.", "B.She may be late for the football game.", "C.She is worried about missing her flight.", "D.She is currently caught in a traffic jam."]
            }  
          ]
        },
        {
          :title => '二、问答题',
          :items => [
            {
              :content => '1. 图中左边有一对平行金属板，两板相距为d，电压为V;两板之间有匀强磁场，磁感应强度大小为B0，方向平行于板面并垂直于纸面朝里。图中右边有一边长为a的正三角形区域EFG(EF边与金属板垂直)，在此区域内及其边界上也有匀强磁场，磁感应强度大小为B，方向垂直于纸面朝里。假设一系列电荷量为q的正离子沿平行于金属板面，垂直于磁场的方向射入金属板之间，沿同一方向射出金属板之间的区域，并经EF边中点H射入磁场区域。不计重力\n
              （1）已知这些离子中的离子${image0}甲到达磁场边界EG后，从边界EF穿出磁场，求离子甲的质量。\n
              （2）已知这些离子中的离子乙从EG边上的I点（图中未画出）穿出磁场，且GI长为 ，求离   子乙的质量。\n
              （3）若这些离子中的最轻离子的质量等于离子甲质量的一半，而离子乙的质量是最大的，问磁场边界上什么区域内可能有离子到达。
              '
            },{
                :content => 'pppppppp_``table:[["摘要_${2}$_", "总账科目_${2}$_", "明细科目_${2}$_", "√_${2}$_", "借款金额__{10}__", "√_${2}$_", "贷方金额__{10}__", "副单据张_${3}$_"], ["{_}", "{_}", "{_}", "{_}", "千", "百", "十", "万", "千", "百", "十", "元", "角", "分", "{_}", "千", "百", "十", "万", "千", "百", "十", "元", "角", "分", "{_}"], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "{_}"]]``_qqqqqqqqqq'
                #:content => 'pppppppp${table:[["摘要_${2}$__@3045@_", "总账科目_${2}$__@1685@_", "明细科目_${2}$__@1624@_", "√_${2}$__@345@_", "借款金额__{10}___@2090@_", "√_${2}$__@345@_", "贷方金额__{10}___@2253@_", "副单据张_${3}$__@345@_"], ["{_}", "{_}", "{_}", "{_}", "千", "百", "十", "万", "千", "百", "十", "元", "角", "分", "{_}", "千", "百", "十", "万", "千", "百", "十", "元", "角", "分", "{_}"], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "{_}"]]}qqqqqqqqqq'
              }
          ]
        }
      ]
    }
   end
end

#describe DocxTemplater::TemplateProcessor do
  #let(:data) { Marshal.load(Marshal.dump(DocxTemplater::TestData::DATA)) } # deep copy
  #let(:base_path) { SPEC_BASE_PATH.join("example_input") }
  #let(:xml) { File.read("#{base_path}/word/document.xml") }
  #let(:parser) { DocxTemplater::TemplateProcessor.new(data) }

  #context "valid xml" do
    #it "should render and still be valid XML" do
      #Nokogiri::XML.parse(xml).should be_xml
      #out = parser.render(xml)
      #Nokogiri::XML.parse(out).should be_xml
    #end

    #it "should accept non-ascii characters" do
      #data[:teacher] = "老师"
      #out = parser.render(xml)
      #out.index("老师").should >= 0
      #Nokogiri::XML.parse(out).should be_xml
    #end

    #it "should escape as necessary invalid xml characters, if told to" do
      #data[:building] = "23rd & A #1 floor"
      #data[:classroom] = "--> 201 <!--"
      #data[:roster][0][:name] = "<#Ai & Bo>"
      #out = parser.render(xml)

      #Nokogiri::XML.parse(out).should be_xml
      #out.index("23rd &amp; A #1 floor").should >= 0
      #out.index("--&gt; 201 &lt;!--").should >= 0
      #out.index("&lt;#Ai &amp; Bo&gt;").should >= 0
    #end

    #context "not escape xml" do
      #let(:parser) { DocxTemplater::TemplateProcessor.new(data, false) }
      #it "does not escape the xml attributes" do
        #data[:building] = "23rd <p>&amp;</p> #1 floor"
        #out = parser.render(xml)
        #Nokogiri::XML.parse(out).should be_xml
        #out.index("23rd <p>&amp;</p> #1 floor").should >= 0
      #end
    #end
  #end

  #context "unmatched begin and end row templates" do
    #it "should not raise" do
      #xml = <<EOF
#<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  #<w:body>
    #<w:tbl>
      #<w:tr><w:tc>
          #<w:p>
            #<w:r><w:t>#BEGIN_ROW:#{:roster.to_s.upcase}#</w:t></w:r>
          #</w:p>
      #</w:tc></w:tr>
      #<w:tr><w:tc>
          #<w:p>
            #<w:r><w:t>#END_ROW:#{:roster.to_s.upcase}#</w:t></w:r>
          #</w:p>
      #</w:tc></w:tr>
      #<w:tr><w:tc>
          #<w:p>
            #<w:r><w:t>#BEGIN_ROW:#{:event_reports.to_s.upcase}#</w:t></w:r>
          #</w:p>
      #</w:tc></w:tr>
      #<w:tr><w:tc>
          #<w:p>
            #<w:r><w:t>#END_ROW:#{:event_reports.to_s.upcase}#</w:t></w:r>
          #</w:p>
      #</w:tc></w:tr>
    #</w:tbl>
  #</w:body>
#</xml>
#EOF
      #expect { parser.render(xml) }.to_not raise_error
    #end

    #it "should raise an exception" do
      #xml = <<EOF
#<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  #<w:body>
    #<w:tbl>
      #<w:tr><w:tc>
          #<w:p>
            #<w:r><w:t>#BEGIN_ROW:#{:roster.to_s.upcase}#</w:t></w:r>
          #</w:p>
      #</w:tc></w:tr>
      #<w:tr><w:tc>
          #<w:p>
            #<w:r><w:t>#END_ROW:#{:roster.to_s.upcase}#</w:t></w:r>
          #</w:p>
      #</w:tc></w:tr>
      #<w:tr><w:tc>
          #<w:p>
            #<w:r><w:t>#BEGIN_ROW:#{:event_reports.to_s.upcase}#</w:t></w:r>
          #</w:p>
      #</w:tc></w:tr>
    #</w:tbl>
  #</w:body>
#</xml>
#EOF
      #expect { parser.render(xml) }.to raise_error(/#END_ROW:EVENT_REPORTS# nil: true/)
    #end
  #end

  #it "should enter no text for a nil value" do
    #xml = <<EOF
#<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
#<w:body>
  #<w:p>Before.$KEY$After</w:p>
#</w:body>
#</xml>
#EOF
    #actual = DocxTemplater::TemplateProcessor.new(:key => nil).render(xml)
    #expected_xml = <<EOF
#<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
#<w:body>
  #<w:p>Before.After</w:p>
#</w:body>
#</xml>
#EOF
    #actual.should == expected_xml
  #end

  #it "should replace all simple keys with values" do
    #non_array_keys = data.reject { |k, v| v.class == Array }
    #non_array_keys.keys.each do |key|
      #xml.index("$#{key.to_s.upcase}$").should >= 0
      #xml.index(data[key].to_s).should be_nil
    #end
    #out = parser.render(xml)

    #non_array_keys.each do |key|
      #out.index("$#{key}$").should be_nil
      #out.index(data[key].to_s).should >= 0
    #end
  #end

  #it "should replace all array keys with values" do
    #xml.index("#BEGIN_ROW:").should >= 0
    #xml.index("#END_ROW:").should >= 0
    #xml.index("$EACH:").should >= 0

    #out = parser.render(xml)

    #out.index("#BEGIN_ROW:").should be_nil
    #out.index("#END_ROW:").should be_nil
    #out.index("$EACH:").should be_nil

    #[:roster, :event_reports].each do |key|
      #data[key].each do |row|
        #row.values.map(&:to_s).each do |row_value|
          #out.index(row_value).should >= 0
        #end
      #end
    #end
  #end

  #it "shold render students names in the same order as the data" do
    #out = parser.render(xml)
    #out.index('Sally').should >= 0
    #out.index('Kumar').should >= 0
    #out.index('Kumar').should > out.index('Sally')
  #end

  #it "shold render event reports names in the same order as the data" do
    #out = parser.render(xml)
    #out.index('Science Museum Field Trip').should >= 0
    #out.index('Wilderness Center Retreat').should >= 0
    #out.index('Wilderness Center Retreat').should > out.index('Science Museum Field Trip')
  #end

  #it "should render 2-line event reports in same order as docx" do
    #event_reports_starting_at = xml.index("#BEGIN_ROW:EVENT_REPORTS#")
    #event_reports_starting_at.should >= 0
    #xml.index("$EACH:NAME$", event_reports_starting_at).should > event_reports_starting_at
    #xml.index("$EACH:NOTES$", event_reports_starting_at).should > event_reports_starting_at
    #xml.index("$EACH:NOTES$", event_reports_starting_at).should > xml.index("$EACH:NAME$", event_reports_starting_at)

    #out = parser.render(xml)
    #out.index('PTA sponsored event. Spoke to Astronaut with HAM radio.').should > out.index('Science Museum Field Trip')
  #end

  #it "should render sums of input data" do
    #xml.index("#SUM").should >= 0
    #out = parser.render(xml)
    #out.index("#SUM").should be_nil
    #out.index("#{data[:roster].count} Students").should >= 0
    #out.index("#{data[:event_reports].count} Events").should >= 0
  #end
#end
