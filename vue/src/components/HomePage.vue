<template>
  <div class="home-page">
    <!-- 浙江锦途 -->
    <section class="home-section">
      <div class="section-header">
        <span class="section-tag">浙江锦途</span>
        <div class="section-line"></div>
      </div>
      <div class="section-intro">
        <p class="intro-text">
          锦途集团是集企业管理咨询、人力资源开发与服务、劳务外包等于一体的综合性人力资源服务企业，是中国领先的专业人力资源服务机构之一，下辖三十余家独立法人机构，锦途先后与安徽、江苏、浙江、湖南、湖北、四川、山西、云南、贵州、河南等省市区多个劳务输出大省的劳动部门建立良好的合作关系，并与全国50多家大中专、技校保持着长期稳定的合作关系，经过多年努力，锦途先后被评为浙江、安徽星级人力资源服务机构。
        </p>
      </div>
    </section>

    <!-- 产品与服务 -->
    <section class="home-section">
      <div class="section-header">
        <span class="section-tag">产品与服务</span>
        <div class="section-line"></div>
      </div>
      <el-carousel height="520px" arrow="always" indicator-position="outside">
        <el-carousel-item>
          <div class="service-card staffing-card">
            <h3 class="service-title">灵活用工</h3>
            <p class="service-desc">
              灵活用工是一种区别于传统全日制固定雇佣关系的用工模式。它强调用工的<span class="highlight">灵活性</span>、<span class="highlight">按需性</span>和<span class="highlight">项目导向性</span>，让企业能够根据业务波动、项目需求或特定技能需要，更加弹性地配置人力资源，同时也为劳动者提供了更多样化的工作选择。
            </p>
            <img class="service-img staffing-img" src="~@/assets/images/img.png" alt="灵活用工" />
          </div>
        </el-carousel-item>
        <el-carousel-item>
          <div class="service-card tax-card">
            <h3 class="service-title">税务筹划</h3>
            <p class="service-desc">
              税务筹划是指在遵守国家税收法律法规的前提下，通过对公司经营、投资、理财等活动的事先规划和安排，合法、合理地<span class="highlight">降低税收负担</span>、<span class="highlight">优化现金流</span>、<span class="highlight">提升经济效益</span>的专业活动。对于企业而言，有效的税务筹划不仅是降低运营成本的关键，更是提升核心竞争力、实现合规稳健发展的重要保障。
            </p>
            <img class="service-img tax-img" src="~@/assets/images/img_1.png" alt="税务筹划" />
          </div>
        </el-carousel-item>
        <el-carousel-item>
          <div class="service-card talent-card">
            <h3 class="service-title">中高级人才搜寻</h3>
            <p class="service-desc">
              中高端人才配置服务为全国各类企业提供专业的人才定向挖掘和全球搜寻服务。拥有完善的后续服务与专家团队，为企业和候选人提供<span class="highlight">上岗人才培训与辅导</span>、<span class="highlight">薪酬调研与薪酬体系设计</span>、<span class="highlight">人才流失防控体系设计与建设</span>、<span class="highlight">企业用人文化塑造与改善</span>等，从而使用人单位以最低的成本发挥所推荐人才的最大作用，并彻底解决人才招募、使用与流失等系列问题。
            </p>
            <img class="service-img talent-img" src="~@/assets/images/img_2.png" alt="中高级人才搜寻" />
          </div>
        </el-carousel-item>
        <el-carousel-item>
          <div class="service-card dispatch-card">
            <h3 class="service-title">劳务派遣服务</h3>
            <p class="service-desc">
              <span class="highlight">劳务派遣</span>是将传统的用工方式中的用工单位和劳动者之间直接的劳动关系进行改变，由浙江锦途（法律意义上的用人单位）与拟派遣的蓝领、金领、普通员工签订劳动合同，将派往实际用工单位从事约定工作，从而使得用工单位更灵活的使用人力资源。
            </p>
            <img class="service-img dispatch-img" src="~@/assets/images/img_3.png" alt="劳务派遣服务" />
          </div>
        </el-carousel-item>
        <el-carousel-item>
          <div class="service-card rpo-card">
            <h3 class="service-title">招聘流程外包</h3>
            <p class="service-desc">
              <span class="highlight">招聘流程外包</span>服务，是指浙江锦途根据企业的用人需要，充分发挥自身招聘会、网站、人才库、渠道输送四大平台优势和专业的招聘团队，采用专业的招聘服务，接受客户单位全部或部分招聘、甄选、培训等工作的委托，提供起点到终点的招聘流程外包服务。
              锦途招聘流程外包服务：针对计划成立新项目团队或分店、分公司的企业，锦途依靠本身的海量的人才储备、庞大的招聘渠道、优秀的寻访能力、专业的甄选方法和经验为客户提供有效的团队建设服务。
            </p>
            <img class="service-img rpo-img" src="~@/assets/images/img_4.png" alt="招聘流程外包" />
          </div>
        </el-carousel-item>
        <el-carousel-item>
          <div class="service-card process-card">
            <h3 class="service-title">生产工序外包</h3>
            <p class="service-desc">
              <span class="highlight">生产工序外包</span>是指企业制造某成品有多道工序，将生产中的非核心某道工序或全部外包给浙江锦途，锦途凭借自身的专业资源和优势按照企业标准，自行招募员工，自主经营管理，接收企业监督，最终以产量为双方结算的依据，与企业进行结算的服务方式。锦途从事的工序外包涉及多行业、多模式；锦途在500强企业中承揽了众多工序外包业务，具备外包服务经验；按相关要求保质保量完成任务，为企业节约大量人工成本。
            </p>
            <img class="service-img process-img" src="~@/assets/images/img_5.png" alt="生产工序外包" />
          </div>
        </el-carousel-item>
      </el-carousel>
    </section>

    <!-- 典型案例 -->
    <section class="home-section">
      <div class="section-header">
        <span class="section-tag">典型案例</span>
        <div class="section-line"></div>
      </div>
      <div class="case-grid">
        <div class="case-card" @click="showCaseDetail('ruili')">
          <img class="case-img" src="~@/assets/images/img_9.png" alt="瑞立集团" />
          <div class="case-name">瑞立集团</div>
        </div>
        <div class="case-card" @click="showCaseDetail('changjiang')">
          <img class="case-img" src="~@/assets/images/img_6.png" alt="长江汽车电子" />
          <div class="case-name">长江汽车电子</div>
        </div>
        <div class="case-card" @click="showCaseDetail('hexing')">
          <img class="case-img" src="~@/assets/images/img_7.png" alt="合兴集团" />
          <div class="case-name">合兴集团</div>
        </div>
        <div class="case-card" @click="showCaseDetail('langshide')">
          <img class="case-img" src="~@/assets/images/img_8.png" alt="朗诗德" />
          <div class="case-name">朗诗德</div>
        </div>
      </div>
    </section>

    <!-- 案例详情弹窗 -->
    <el-dialog
      title=""
      v-model="caseDialogVisible"
      width="800px"
      top="5vh"
      :close-on-click-modal="true"
      custom-class="case-dialog"
    >
      <div class="case-detail" v-if="currentCase === 'ruili'">
        <h2 class="case-detail-title">化解500人用工荒 实现逆势增长</h2>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">客户简介：</h4>
          <p class="case-detail-text">
            瑞立集团，创办于1987年，是专业生产汽车零部件的企业。瑞立生产的产品已为中国一汽集团、东风集团、上汽集团等30多家国内汽车制造厂以及欧洲、美洲和日本车系等提供维修服务，并逐步进入了全球汽车配件采购体系，并为2008北京奥运会公交车、国庆60周年阅兵军车、2010年上海世博会、广州亚运会新能源客车提供产品配套服务。瑞立已经发展成为温州地区汽摩配行业龙头企业、全国民营企业500强、浙江省制造业百强、浙江省工业行业龙头骨干企业。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">项目背景：</h4>
          <p class="case-detail-text">
            2024年末至2025年初，受行业周期性及春节返乡潮影响，瑞立公司的核心生产基地遭遇了严峻挑战：大量一线员工集中离职返乡，导致其机加数控、产品装配等岗位，骤然出现近500人的巨大用工缺口。这一缺口不仅出现在生产高峰前夕，更直接威胁到订单交付周期、生产线稳定运行以及全年产能目标的达成。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">解决方案：驻场洞察 + 敏捷响应 + 精准交付</h4>
          <p class="case-detail-text">
            作为瑞立集团长期信赖的人力资源战略合作伙伴，浙江锦途人力资源有限公司的驻场服务团队第一时间捕捉到了这一异常波动。驻场顾问凭借其深度融入客户现场的优势，迅速、精准地将用工缺口规模、岗位技能要求、紧急程度等关键信息反馈至锦途项目管理层。
          </p>
          <p class="case-detail-text">
            基于前期战略合作中建立的信任与对锦途响应能力的认可，瑞立集团紧急授权锦途团队全权负责该紧急用工缺口，要求短时间内实现缺口填补。
          </p>
          <p class="case-detail-text">
            锦途团队迅速启动应急响应机制：
          </p>
          <p class="case-detail-text">
            <strong>精准资源匹配：</strong>利用覆盖全国的庞大人才数据库及250+就业市场深度合作关系，从安徽、河南、四川等劳务输出大省定向筛选数控操作、装配等技能匹配的候选人员。
          </p>
          <p class="case-detail-text">
            <strong>快速输送通道：</strong>依托长期合作的职业院校与技工学校资源，开辟"面试-体检-入职-住宿"一站式绿色通道，实现从人员储备到上岗的无缝衔接。
          </p>
          <p class="case-detail-text">
            <strong>驻场同步管理：</strong>驻场顾问与生产现场管理层同步对接，实时掌握产线排产与人员调配需求，灵活调整人员配置、优化人力效能。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">项目成果：</h4>
          <p class="case-detail-text">
            锦途团队在预定时间内完成500人用工补充，保障了生产线24小时满负荷运转，订单如期交付，全年营收目标得以顺利实现。
          </p>
          <p class="case-detail-text">
            基于本次成功交付，瑞立集团与浙江锦途签署了长期战略合作协议，将锦途列为其首选人力资源服务供应商，并授予年度"优秀战略合作伙伴"称号。
          </p>
        </div>
      </div>

      <div class="case-detail" v-if="currentCase === 'changjiang'">
        <h2 class="case-detail-title">闪电集结 质效双赢</h2>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">客户简介：</h4>
          <p class="case-detail-text">
            浙江长江汽车电子有限公司成立于2004年12月10日，注册地位于温州经济技术开发区，是长江汽车电器集团旗下的高新技术企业。公司专注于智能座舱系统、新能源汽车电子系统等领域的研发制造，为德国大众、特斯拉等国际车企提供产品与服务，2023年实现年产值超50亿元。企业构建了覆盖温沪德美的全球研发网络，拥有1500余人的技术团队。截至2023年，公司参保员工达6122人，车间配备现代化生产设施，宿舍采用中央空调等配套体系。在政企合作方面，2023年成为龙湾区复工复产安全示范单位。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">项目背景：</h4>
          <p class="case-detail-text">
            2022年底，随着汽车消费市场的强劲复苏及订单激增，浙江长江汽车电子有限公司面临巨大的生产交付压力。其核心产线——负责生产汽车中控控制器及车窗玻璃升降开关的部门，因业务扩张和季节性因素，出现了大规模的一线操作工缺口，人员需求迫在眉睫。
          </p>
          <p class="case-detail-text">
            <strong>1. 极短时限：</strong>要求在短短10天内，完成300名合格一线操作工的招聘及全套入职流程（包含面试、筛选、体检、合同签订、培训、上岗）。
          </p>
          <p class="case-detail-text">
            <strong>2. 质量保障：</strong>在追求速度的同时，绝不降低人才标准，必须确保入职人员符合岗位要求。
          </p>
          <p class="case-detail-text">
            <strong>3. 岗位要求高：</strong>
          </p>
          <p class="case-detail-text">
            · 工作强度大：以两班倒站班为主，对员工体能和耐力有较高要求。
          </p>
          <p class="case-detail-text">
            · 专业规范严：必须穿着静电衣作业，严格遵守电子车间无尘防静电规范，保障精密汽车电子元器件（中控控制器、升降开关）的生产质量。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">解决方案：</h4>
          <p class="case-detail-text">
            <strong>1.精准定位，多渠道闪电启动：</strong>
          </p>
          <p class="case-detail-text">
            · 深度需求分析：项目组第一时间与长江汽车电子HR及产线负责人进行详细对接，精准把握岗位核心要求（体能、纪律性、防静电意识）、工作环境特点及公司文化。
          </p>
          <p class="case-detail-text">
            · 全渠道饱和式覆盖：在极短时间内，同步启动线上（主流招聘网站、地方性人才网、供应商合作商）、线下（深入目标劳动力富集区域设点、老员工推荐激励、合作院校/机构推荐等多维度招聘渠道，形成强大的信息覆盖网。招聘信息突出强调"正规大厂"、"待遇明确"、"快速入职"等关键吸引力。
          </p>
          <p class="case-detail-text">
            · 预筛选前置：在简历投递和初步沟通阶段即设置清晰门槛（年龄、体能、工作意愿、理解能力），大幅提升后续环节效率。
          </p>
          <p class="case-detail-text">
            <strong>2.严控流程，质量速度并行：</strong>
          </p>
          <p class="case-detail-text">
            · 高效集中面试：组织多场次、高密度的集中面试/初筛，由经验丰富的招聘官和产线代表共同把关，重点考察应聘者的稳定性、吃苦耐劳精神以及对站班、静电防护要求的理解和接受度。
          </p>
          <p class="case-detail-text">
            · "一站式"入职服务：优化整合入职流程。安排专车接送面试/体检；协调定点医院快速体检通道；设立临时入职办理点，现场完成合同签订、资料收集、基础信息录入；同步启动线上入职培训模块。
          </p>
          <p class="case-detail-text">
            · 强化岗前适应性培训：虽然时间紧迫，但浙江锦途人力坚持安排紧凑且实用的岗前培训。核心内容包括：公司文化制度宣导、产线安全规范（重中之重）、静电防护知识与静电衣穿戴规范演练、基础质量意识、两班倒作息适应指引。确保人员上岗即能基本遵守规范，降低因不了解而产生的流失。
          </p>
          <p class="case-detail-text">
            <strong>3.风险预控，保障留存基础：</strong>
          </p>
          <p class="case-detail-text">
            · 透明沟通：在招聘全过程中，向候选人清晰、真实地传递工作内容、强度、环境及待遇，避免因信息不对称导致入职后快速离职。
          </p>
          <p class="case-detail-text">
            · 关注体验：从应聘者接触的第一刻起，注重服务体验，专人答疑解惑，营造被尊重和重视的氛围。
          </p>
          <p class="case-detail-text">
            · 入职关怀：入职初期安排驻厂管理人员，及时解决新员工遇到的生活和工作困难，增强归属感。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">项目成果与价值：</h4>
          <p class="case-detail-text">
            · 速度奇迹：在客户要求的10天期限内，浙江锦途人力仅用了<strong>7天</strong>就完成了<strong>300人</strong>的招聘及入职流程。
          </p>
          <p class="case-detail-text">
            · 质量保障：入职人员均通过严格筛选，基本满足岗位技能、体能及规范性要求。上岗首日人员到位率高达98%。
          </p>
          <p class="case-detail-text">
            · 稳定达标：项目启动后，在浙江锦途人力的持续跟进和与长江汽车电子的协同管理下，该批次员工的后续在岗率成功稳定在40%以上，有效保障了客户产线的平稳运行和订单交付。
          </p>
          <p class="case-detail-text">
            · 客户高度认可：浙江锦途人力在极限时间压力下展现出的专业能力、应急响应速度和结果交付能力，赢得了长江汽车电子管理层的高度评价和信任，进一步巩固了双方的战略合作伙伴关系。
          </p>
        </div>
      </div>

      <div class="case-detail" v-if="currentCase === 'hexing'">
        <h2 class="case-detail-title">精准赋能 长期共赢</h2>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">客户简介：</h4>
          <p class="case-detail-text">
            合兴汽车电子股份有限公司——成立于2006年，在电子连接器制造领域实力雄厚，总部位于浙江省乐清市虹桥镇，入选"工信部智能制造试点示范"项目。公司现阶段主要生产汽车连接器，是国内少数进入全球汽车零部件采购体系的本土企业之一，为大众、宝马、博格华纳等国际企业供货，是温州地区重点高新技术及智能制造标杆企业。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">项目背景：</h4>
          <p class="case-detail-text">
            核心连接器产量大、需求旺盛，产品有大量缺口。合兴汽车电子生产的核心连接器应用于汽车的各个核心部件，市场需求旺盛，订单量持续增长。为保障交付，需要对大量生产一线员工及时补充。
          </p>
          <p class="case-detail-text">
            · 用人规模大：随着新产线陆续投产和产能提升，合兴多条产线需大量且持续补充操作员，用工需求常年稳定。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">解决方案：</h4>
          <p class="case-detail-text">
            <strong>1. 驻厂服务，需求前置预判：</strong>锦途人力为该客户配备驻厂人员，与客户用人需求部门实时对接，匹配生产节奏和排产计划提前启动招聘，避免用工延迟。
          </p>
          <p class="case-detail-text">
            <strong>2. 人才筛选，满足用工高要求：</strong>电子连接器生产对操作精密度和产品洁净度要求极高，且大多数岗位采用站立作业模式。驻厂人员深入了解产线细分岗位对人的特殊要求，在招聘前端进行精准画像，确保人才筛选中重点关注候选人的站班适应性和手工作业细心度。
          </p>
          <p class="case-detail-text">
            <strong>3. 人员储备，确保持续稳定供应：</strong>锦途人力不仅为当前缺口招聘，当客户没有需求时，锦途也会提前物色适岗候选人，为可能到来的下一波用工简提前储备，确保人员"随需随到"。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">项目成果与价值：</h4>
          <p class="case-detail-text">
            · 8年的长期合作：锦途自合作以来始终稳定保障生产运营。
          </p>
          <p class="case-detail-text">
            · 多元化的人才供给：提供正式工、临时工、项目工等灵活多样的用工模式。
          </p>
          <p class="case-detail-text">
            · 灵活机动、快速响应的合作机制：根据订单波动弹性调整人力配置，提高交付效率和降低用工成本。
          </p>
          <p class="case-detail-text">
            · 保障区域用工成本优势：锦途深度人力市场资源，帮助客户在薪资源结构中保持竞争力。
          </p>
        </div>
      </div>

      <div class="case-detail" v-if="currentCase === 'langshide'">
        <h2 class="case-detail-title">急招攻坚 速度制胜</h2>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">客户简介：</h4>
          <p class="case-detail-text">
            浙江朗诗德健康饮水设备股份有限公司，创立于2009年，总部位于浙江省温州，是一家集研发、生产、销售于一体的行业领军企业。公司以"水生态健康"为核心理念，致力于为用户提供高品质的净饮水解决方案。其产品线覆盖家用、商用全场景净饮水设备，并跻身"国家高新技术企业"行列。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">项目背景：</h4>
          <p class="case-detail-text">
            2025年春节后，郎诗德面临前所未有的用工压力。年前外省员工大量返乡且返岗率极低，导致多条生产线因人手严重不足无法正常运转。与此同时，国内外市场订单迎来高峰期，如果在短时间内无法补充足够的生产人员，将面临订单超期违约和市场份额流失的巨大风险。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">解决方案：</h4>
          <p class="case-detail-text">
            浙江锦途人力接到客户紧急求助后，迅速启动应急预案，以"找人"和"留人"为核心，聚焦以下关键行动：
          </p>
          <p class="case-detail-text">
            <strong>1. 多渠道并进，广纳人员：</strong>启动全国核心人力输出基地、职业学院优质合作院校招聘网络，开启线上流量招聘发力，大幅度扩大信息覆盖范围，在竞争激烈的人力市场抢占先机。
          </p>
          <p class="case-detail-text">
            <strong>2. 优化流程 提升转化：</strong>极简招聘流程，安排专车接送，高效运作签约。
          </p>
          <p class="case-detail-text">
            <strong>3. 关注体验，保障留存：</strong>锦途为郎诗德设立了驻厂服务点，安排专人驻场服务，提供从工作到生活的关怀，帮助新员工快速融入，有效提升入职稳定率。
          </p>
        </div>

        <div class="case-detail-section">
          <h4 class="case-detail-subtitle">项目成果与价值：</h4>
          <p class="case-detail-text">
            · 高效交付：通过高效的反应机制和体系化招聘网络，成功为企业注入大量适岗一线员工，保障了生产旺季的订单交付。
          </p>
          <p class="case-detail-text">
            · 人员稳定：精细化的人员筛选和驻厂关怀显著提升了员工留存率，为企业降低了重复招聘的成本。
          </p>
          <p class="case-detail-text">
            · 客户认可：锦途展现出的专业性和高效性，获得了郎诗德高层的高度评价，锦途被列为重要的人力资源战略合作伙伴。
          </p>
        </div>
      </div>
    </el-dialog>

    <!-- 服务优势 -->
    <section class="home-section">
      <div class="section-header">
        <span class="section-tag">服务优势</span>
        <div class="section-line"></div>
      </div>
      <div class="advantage-intro">
        <div class="advantage-block">
          <h4 class="advantage-block-title">🔍 全方位服务能力与灵活定制化方案</h4>
          <div class="advantage-item">
            <p class="advantage-item-title">· 一站式人力资源解决方案</p>
            <p class="advantage-item-text">锦途人力提供涵盖招聘外包、劳务派遣、岗位工序外包、HR管理咨询、培训开发、薪酬绩效管理等全链条服务，能够根据企业需求定制"从员工招聘到解聘"的全流程服务，帮助客户优化人力资源配置、降低成本。</p>
          </div>
          <div class="advantage-item">
            <p class="advantage-item-title">· 高度灵活的定制化服务</p>
            <p class="advantage-item-text">公司以"量身定制"为核心竞争力，针对不同行业（如制造业、汽车零部件等）和企业规模设计差异化方案，快速响应市场变化与用工形式多样化需求。</p>
          </div>
        </div>
        <div class="advantage-block">
          <h4 class="advantage-block-title">🌐 庞大的资源网络与高效输送能力</h4>
          <div class="advantage-item">
            <p class="advantage-item-title">· 全国性人才输送网络</p>
            <p class="advantage-item-text">已与全国250多个就业市场建立深度合作，掌握庞大的产业工人与专业技术人才数据库，2024年仅向瑞安本地企业输送一线产业工人即达3万人次。</p>
          </div>
          <div class="advantage-item">
            <p class="advantage-item-title">· 精准匹配与高效执行</p>
            <p class="advantage-item-text">依托数据化资源池和专业化团队，实现人才需求的高效匹配。简历处理平均用时仅2小时，及时处理率超90%，显著提升企业用工效率。</p>
          </div>
        </div>
        <div class="advantage-block">
          <h4 class="advantage-block-title">🏢 区域战略布局与政策平台支持</h4>
          <div class="advantage-item">
            <p class="advantage-item-title">· 核心产业园区位优势</p>
            <p class="advantage-item-text">公司总部位于浙江温州人力资源服务产业园（瑞安），该产业园是浙南地区首个专业化人力资源服务平台，入驻企业享受政策、资源及品牌集聚红利。2025年，锦途人力进一步入驻温州鹿城省级人力资源产业园，强化区域影响力。</p>
          </div>
          <div class="advantage-item">
            <p class="advantage-item-title">· 借力区域产业升级需求</p>
            <p class="advantage-item-text">深度服务温州地区的民营经济，尤其针对汽车零部件等主导产业提供人才支持，与地方经济转型形成协同效应。</p>
          </div>
        </div>
      </div>
      <div class="advantage-row">
        <div class="advantage-col">
          <h4 class="advantage-col-title">🏆 行业认可</h4>
          <div class="advantage-col-item">
            <p class="advantage-col-item-title">· 资质与信誉保障</p>
            <p class="advantage-col-text">2024年被评为"浙江省企业信用等级四星级单位"，并入选省级人力资源服务产业园重点企业，体现其在合规性、服务质量上的权威认可。</p>
          </div>
          <img class="advantage-cert-img" src="~@/assets/images/certificate-4star.png" alt="企业信用等级四星级证书" />
        </div>
        <div class="advantage-col">
          <div class="advantage-col-item">
            <p class="advantage-col-item-title">· 深获业界信赖</p>
            <p class="advantage-col-text">锦途人力在人力资源服务领域深耕细作，专业能力与卓越服务赢得了区域内企业客户的高度认可与信赖。最有力的证明，便是我们获得合作客户颁发的荣誉授牌，例如"优秀战略合作伙伴"和"优秀供应商"等称号。</p>
          </div>
          <img class="advantage-cert-img" src="~@/assets/images/partner.png" alt="优秀合作伙伴" />
        </div>
      </div>
    </section>

    <!-- 企业资质 -->
    <section class="home-section">
      <div class="section-header">
        <span class="section-tag">企业资质</span>
        <div class="section-line"></div>
      </div>
      <div class="qualification-grid">
        <div
          class="qualification-card"
          v-for="(img, idx) in qualificationImages"
          :key="idx"
          @click="previewImage(img)"
        >
          <img :src="img" class="qualification-img" loading="lazy" alt="企业资质" />
        </div>
      </div>
    </section>

    <!-- 资质图片预览弹窗 -->
    <el-dialog
      v-model="previewVisible"
      width="30%"
      top="5vh"
      :close-on-click-modal="true"
      :show-close="true"
      custom-class="preview-dialog"
    >
      <img :src="previewSrc" class="preview-full-img" alt="预览" />
    </el-dialog>

    <!-- 企业荣誉 -->
    <section class="home-section">
      <div class="section-header">
        <span class="section-tag">企业荣誉</span>
        <div class="section-line"></div>
      </div>
      <div class="honor-scroll-wrapper">
        <div class="honor-scroll-track">
          <img
            v-for="(img, idx) in honorImages"
            :key="'a' + idx"
            :src="img"
            class="honor-img"
            loading="lazy"
            alt="企业荣誉"
          />
          <!-- 复制一份实现无缝循环 -->
          <img
            v-for="(img, idx) in honorImages"
            :key="'b' + idx"
            :src="img"
            class="honor-img"
            loading="lazy"
            alt="企业荣誉"
          />
        </div>
      </div>
    </section>
  </div>
</template>

<script>
import '@/styles/components/HomePage.css'

export default {
  name: 'HomePage',
  data() {
    return {
      caseDialogVisible: false,
      currentCase: '',
      previewVisible: false,
      previewSrc: ''
    }
  },
  computed: {
    qualificationImages() {
      const ctx = require.context('@/assets/企业资质', false, /\.(png|jpg|jpeg)$/i)
      return ctx.keys()
        .sort((a, b) => {
          const nA = parseInt(a.match(/(\d+)/)[1])
          const nB = parseInt(b.match(/(\d+)/)[1])
          return nA - nB
        })
        .map(key => ctx(key))
    },
    honorImages() {
      const ctx = require.context('@/assets/企业荣誉', false, /\.(png|jpg|jpeg)$/i)
      return ctx.keys()
        .sort((a, b) => {
          // 数字优先排序，相同数字再按扩展名排序（png 优先）
          const nA = parseInt(a.match(/(\d+)/)[1])
          const nB = parseInt(b.match(/(\d+)/)[1])
          if (nA !== nB) return nA - nB
          return a.localeCompare(b)
        })
        .map(key => ctx(key))
    }
  },
  methods: {
    showCaseDetail(caseName) {
      this.currentCase = caseName
      this.caseDialogVisible = true
    },
    previewImage(src) {
      this.previewSrc = src
      this.previewVisible = true
    }
  }
}
</script>

<style>
/* 弹窗挂载在 body 下，需全局样式 */
.case-dialog .el-dialog__header {
  padding: 0;
  border-bottom: none;
}

.case-dialog .el-dialog__body {
  padding: 24px 32px 32px;
}

/* 资质预览弹窗 */
.preview-dialog .el-dialog__header {
  padding: 8px 16px 0;
  border-bottom: none;
}

.preview-dialog .el-dialog__body {
  padding: 16px;
}

.preview-full-img {
  width: 100%;
  display: block;
  border-radius: 4px;
}
</style>
