import paramiko
import re

# 서버 메모리 값 파싱 테스트
def get_used_memory(server_ip, username, password):    
    try:
        # SSH 클라이언트 생성
        client = paramiko.SSHClient()

        # 호스트 키 자동으로 추가 (처음 접속 시에만 필요)
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

        # SSH 서버에 연결
        client.connect(server_ip, username=username, password=password)

        # free -h 명령을 실행하고 출력 결과를 가져옴
        stdout = client.exec_command("free -h")
        result_str = stdout.read().decode("utf-8")

        # 출력 결과에서 사용 가능한 메모리 정보를 파싱
        lines = result_str.split('\n')
        for line in lines:
            if "Mem:" in line:
                parts = re.findall(r'\d+', line)
                total_memory = int(parts[0])
                used_memory = int(parts[1])
                return used_memory, total_memory

    except Exception as e:
        print(f"Error: {e}")

    finally:
        # SSH 세션 종료
        client.close()

if __name__ == "__main__":
    server_ip = "192.168.11.11"
    username = "secret"
    password = "secret"

    result = get_used_memory(server_ip, username, password)

    if result:
        used_memory, total_memory = result
        print(f"Total Memory: {total_memory} MB")
        print(f"Used Memory: {used_memory} MB")
    else:
        print("Failed to retrieve memory information.")