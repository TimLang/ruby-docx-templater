# -*- encoding : utf-8 -*-
require 'spec_helper'

module DocxTemplater
  module TestData
    DATA = {
      :subject => '高考模拟卷一',
      #:teacher => '土豪金',
      :teacher => '土豪金高端大气上档次哈哈',
      :questions => [
        {
          :title => '一、选择题（本题包括21小题，每小题给出的四个选项中，有的只有一个选项正确，有的有多个选项正确，全部选对的得6分，选对但不全的得3分，有选错的得0分）',
          :items => [
            {
              :content => '1. 人在恐惧、紧张时，在内脏神经的支配下，肾上腺髓质释放的肾上腺素增多，该激素可用于心脏，使心率加快。下列叙述错误的是${image2}（    ）',
              :choice => [
                '  A．该肾上腺素作用的靶器官包括心脏',
                '  B．该实例包含神经调节和体液调节',
                '  C．该肾上腺素通过神经纤维运输到心脏',
                '  D．该实例中反射弧是实现神经调节的结构基础'
              ]
            },
            {
              :content => '2. 番茄幼苗在缺镁的培养液中培养一段时间后，与对照组相比，其叶片光合作用强度下降，原因是（    ）',
              :choice => [
                '  A．光反应强度升高，暗反应强度降低',
                '  B．光反应强度降低，暗反应强度降低',
                '  C．反应强度不变，暗反应强度降低',
                '  D．反应强度降低，暗反应强度不变'
              ]
            }
          ]
        },
        {
          :title => '二、问答题',
          :items => [
            {
              :content => '1. 图中左边有一对平行金属板，两板相距为d，电压为V;两板之间有匀强磁场，磁感应强度大小为B0，方向平行于板面并垂直于纸面朝里。图中右边有一边长为a的正三角形区域EFG(EF边与金属板垂直)，在此区域内及其边界上也有匀强磁场，磁感应强度大小为B，方向垂直于纸面朝里。假设一系列电荷量为q的正离子沿平行于金属板面，垂直于磁场的方向射入金属板之间，沿同一方向射出金属板之间的区域，并经EF边中点H射入磁场区域。不计重力${newline}
              （1）已知这些离子中的离子${image0}甲到达磁场边界EG后，从边界EF穿出磁场，求离子甲的质量。${newline}
              （2）已知这些离子中的离子乙从EG边上的I点（图中未画出）穿出磁场，且GI长为 ，求离   子乙的质量。${newline}
              （3）若这些离子中的最轻离子的质量等于离子甲质量的一半，而离子乙的质量是最大的，问磁场边界上什么区域内可能有离子到达。
              '
            },{
                :content => '2．【生物——选修1：生物技术实践】(8分) 为了探究6―BA和IAA对某些菊花品种茎尖外植物再生丛芽的影响，某研究小组在MS培养基中加入6―BA和IAA，配制成四种培养基(见下表)，灭菌后分别接种数量相同、生长状态一致、消毒后地的尖外植体，在适宜条件下培养一段时间后，统计再生丛芽外植体的比率(m)，以及再生丛芽外植体上的丛芽平均数(n)，结果如下表。${newline}
                (2)在该实验中，自变量是       ，因变量是        ，自变量的取值范围是            。${newline}
                (3)从实验结果可知，诱导从芽总数量少的培养基是          号培养基。${newline}
                (4)为了诱导该菊花试管菌生根，培养基中一般不加入          。(填“6―BA”或“IAA”)。'
              }
          ]
        }
      ]
    }
    #DATA = {
        #:teacher => "Priya Vora",
        #:building => "Building #14",
        #:classroom => :'Rm 202',
        #:district => "Washington County Public Schools",
        #:senority => 12.25,
        #:roster => [
            #{:name => 'Sally', :age => 12, :attendence => '100%'},
            #{:name => :Xiao, :age => 10, :attendence => '94%'},
            #{:name => 'Bryan', :age => 13, :attendence => '100%'},
            #{:name => 'Larry', :age => 11, :attendence => '90%'},
            #{:name => 'Kumar', :age => 12, :attendence => '76%'},
            #{:name => 'Amber', :age => 11, :attendence => '100%'},
            #{:name => 'Isaiah', :age => 12, :attendence => '89%'},
            #{:name => 'Omar', :age => 12, :attendence => '99%'},
            #{:name => 'Xi', :age => 11, :attendence => '20%'},
            #{:name => 'Noushin', :age => 12, :attendence => '100%'}
        #],
        #:event_reports => [
            #{:name => 'Science Museum Field Trip', :notes => 'PTA sponsored event. Spoke to Astronaut with HAM radio.'},
            #{:name => 'Wilderness Center Retreat', :notes => '2 days hiking for charity:water fundraiser, $10,200 raised.'}
        #],
        #:created_at => "11-12-03 02:01"
    #}
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
