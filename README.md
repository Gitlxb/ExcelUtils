# 手续费拆分工具 - Spring Boot + Vue 项目

## 项目结构

```
Project/
├── src/main/java/          # Spring Boot后端
├── src/main/resources/     # 配置文件
├── vue/                    # Vue前端
│   ├── public/
│   └── src/
└── uploads/                # 文件上传目录
```

## 启动步骤

### 1. 启动后端（Spring Boot）

```bash
cd D:\Project
mvn spring-boot:run
```

后端将运行在：http://localhost:8080

### 2. 启动前端（Vue）

```bash
cd D:\Project\vue
npm install
npm run serve
```

前端将运行在：http://localhost:8081

## 使用说明

1. 访问 http://localhost:8081
2. 拖拽或点击上传 .xlsx 文件
3. 等待处理完成
4. 下载处理后的文件

## API接口

- **POST** `/api/shouxufei/upload` - 上传并处理Excel文件
- **GET** `/api/shouxufei/download/{fileName}` - 下载处理后的文件

## 功能特性

✅ 自动识别并应用6种条件规则  
✅ 计算手续费（E列）和实际金额（F列）  
✅ 按G列内容拆分成多个工作表  
✅ 自动添加合计行  
✅ 特殊条件行标黄标记  
✅ J、K列重复值标记  
✅ 账号信息汇总  

## 技术栈

**后端：**
- Spring Boot 3.5.14
- Apache POI 5.2.5
- Java 17

**前端：**
- Vue 3
- Element Plus
- Axios
