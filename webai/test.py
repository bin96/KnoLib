import socket

def check_tcp_connection(ip_address, port, timeout=2):
    """
    检查指定IP和端口的TCP连接是否通
    
    参数:
        ip_address: 要检查的IP地址
        port: 要检查的端口号
        timeout: 连接超时时间(秒)，默认2秒
    
    返回:
        如果连接成功，返回True，否则返回False
    """
    try:
        # 创建一个TCP socket
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            # 设置超时时间
            s.settimeout(timeout)
            
            # 尝试连接到指定的IP和端口
            result = s.connect_ex((ip_address, port))
            
            # connect_ex返回0表示连接成功，其他值表示失败
            return result == 0
    
    except socket.error as e:
        print(f"Socket error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")
    
    return False

# 示例用法
if __name__ == "__main__":
    ip = "192.168.96.29"  # 本地回环地址
    port = 11435          # HTTP端口
    
    if check_tcp_connection(ip, port):
        print(f"TCP连接到{ip}:{port}成功!")
    else:
        print(f"TCP连接到{ip}:{port}失败!")