# mes_terminal
MES扫码外挂终端，可通过配置将现场扫码记录上传到MES系统，也可用于喷码机不直接对接时，喷码数据扫码上传

## 功能说明

### 1. 登录/登出功能
- 输入用户名、密码和设备编码进行登录
- 登录成功后自动获取 Token（有效期 24 小时）
- 系统会自动每 23 小时调用延期接口刷新 Token
- 点击"登出"按钮可退出系统

### 2. 工单选择
- 登录后自动获取当前设备绑定的工单列表
- 通过下拉列表选择要使用的工单
- 点击"刷新"按钮可重新获取工单列表

### 3. 扫码功能
- 选择工单后，点击"开始扫码"按钮进入扫码状态
- 使用扫码枪扫描条码（扫码枪需设置为键盘模式，扫描后自动回车）
- 每次扫码成功会自动计数并播放提示音
- 扫码失败时弹窗提示错误信息并停止扫码

### 4. 离线模式
- 勾选"离线模式"复选框后启用
- 离线模式下扫码数据保存在本地 SQLite 数据库（offline_data.db）
- 切换到在线模式后，点击"上传离线数据"可将缓存的数据上传到服务器
- 状态栏显示待上传数据数量

### 5. 其他功能
- **扫码计数**: 实时显示当前扫码数量
- **声音提示**: 扫码成功/失败播放不同提示音
- **导出记录**: 可将所有扫码记录导出为 CSV 文件
- **操作日志**: 记录所有操作历史

### 6. 系统设置（管理员功能）
- **API 接口配置**: 通过菜单栏"系统" -> "API 设置"访问
  - 需要输入管理员密码验证（默认密码：123456）
  - 可配置基础 URL、各接口路径、超时时间等
  - 支持恢复默认设置
- **修改管理员密码**: 通过菜单栏"系统" -> "修改管理员密码"访问

## 安装步骤

1. 确保已安装 Python 3.7+

2. 安装依赖：
```bash
pip install -r requirements.txt
```

3. 运行程序：
```bash
python mes_terminal.py
```

## 配置文件说明

程序首次运行时会自动生成 `config.ini` 配置文件，内容如下：

```ini
[api]
base_url = http://172.16.0.10:8080/Mes/
login_path = /Device/Login
prolong_path = /Devicer/reLogin
crossing_path = /Device/Cross_station
mmo_list_path = /Device/get_MmoList

[admin]
password = 123456

[settings]
timeout = 10
auto_prolong_hours = 23


```

## 界面说明

### 菜单栏
- **系统**
  - API 设置（需管理员密码验证）
  - 修改管理员密码
  - 退出
- **帮助**
  - 关于

### 主界面
<img width="902" height="752" alt="1" src="https://github.com/user-attachments/assets/070d89de-d2e2-4bfb-9621-8b048482a713" />
<img width="902" height="752" alt="2" src="https://github.com/user-attachments/assets/dc3eb89e-f7da-43d6-8ecc-e741e12ca454" />
<img width="902" height="752" alt="login" src="https://github.com/user-attachments/assets/14f37cda-e0fb-4b2d-8ea3-f797b91138f0" />


## API 接口说明

### 默认配置，完成接口由基础URL和各接口拼接而成
- 基础 URL: `http://172.16.10.250:8911/api/Mes/`
- 登录接口：(POST)
- 延期接口：(POST)
- 工单列表：(GET)
- 出站接口：(POST)

### 请求/响应格式

**登录接口**
```json
// 请求
{
    "userName": "JieKou01",
    "password": "JieKou",
    "deviceCode": "P10-10"
}

// 响应
{
    "state": 200,
    "data": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
}
```

**出站接口**
```json
// 请求
{
    "deviceCode": "P10-10",
    "mmoCode": "MMO20240101-001",
    "labels": [
        {
            "label": "SN123456",
            "qty": 1,
            "result": 10
        }
    ]
}

// 响应
{
    "state": 200,
    "exId": 0,
    "code": "success",
    "msg": "success",
    "data": null
}
```

## 注意事项

1. **扫码枪设置**: 需要设置为键盘模式，扫描后自动发送回车键
2. **离线数据**: 离线模式下的数据在上传前保存在本地数据库文件中
3. **Token 有效期**: Token 有效期为 24 小时，系统会自动延期，但关闭程序后重新登录需要重新获取 Token
4. **扫码失败**: 扫码失败时系统会停止扫码，需要用户确认后才能继续
5. **管理员密码**: 默认密码为 123456，建议首次使用后修改

## 常见问题

**Q: 扫码后没有反应？**
A: 检查是否已登录、是否选择了工单、是否点击了"开始扫码"按钮

**Q: 离线数据如何上传？**
A: 登录后点击"上传离线数据"按钮，系统会自动将所有未上传的数据上传到服务器

**Q: 如何查看历史扫码记录？**
A: 点击"导出记录"按钮，可将所有记录导出为 CSV 文件查看

**Q: 如何修改 API 接口地址？**
A: 点击菜单栏"系统" -> "API 设置"，输入管理员密码验证后即可修改

**Q: 管理员密码忘了怎么办？**
A: 直接编辑 `config.ini` 文件，修改 `[admin]`  section 下的 `password` 值

**Q: 如何修改自动延期时间？**
A: 在"API 设置"中修改"自动延期间隔（小时）"，建议设置为 23 小时（小于 Token 有效期 24 小时）
