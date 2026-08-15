"""RSA 授权文件生成器 —— 开发者使用"""
import os
import sys
import json

# 从 sferzaibo 导入签名函数
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sferzaibo import rsa_sign, get_machine_id


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    private_key_path = os.path.join(script_dir, "private_key.pem")

    if not os.path.exists(private_key_path):
        print(f"错误: 找不到私钥文件 {private_key_path}")
        sys.exit(1)

    print("=" * 50)
    print("  RSA 授权文件生成器")
    print("=" * 50)
    print(f"私钥: {private_key_path}")
    print()

    # 交互输入
    customer = input("客户名称: ").strip()
    if not customer:
        print("客户名称不能为空")
        sys.exit(1)

    mid = input("机器码 (回车=any 不绑定机器): ").strip()
    if not mid:
        mid = "any"
    elif mid.lower() == "this":
        mid = get_machine_id()
        print(f"  本机机器码: {mid}")

    exp_input = input("有效天数 (回车=never 永不过期): ").strip()
    from datetime import date, timedelta
    issue = date.today()
    if not exp_input:
        exp = "never"
    else:
        days = int(exp_input)
        exp_date = issue + timedelta(days=days)
        exp = exp_date.isoformat()
        print(f"  到期日期: {exp} (签发 {issue.isoformat()} + {days}天)")

    issue_str = issue.isoformat()

    # 构建数据
    data = {
        "customer": customer,
        "machine_id": mid,
        "expire_date": exp,
        "issue_date": issue_str,
    }
    message = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    signature = rsa_sign(message, private_key_path)

    lic = {
        "data": data,
        "signature": signature,
    }

    output_path = os.path.join(script_dir, "license.lic")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(lic, f, ensure_ascii=False, indent=2)

    print()
    print(f"授权文件已生成: {output_path}")
    print(f"  客户: {customer}")
    print(f"  机器码: {mid}")
    print(f"  过期日期: {exp}")
    print()
    print("将 license.lic 发放给用户即可。")


if __name__ == "__main__":
    main()
