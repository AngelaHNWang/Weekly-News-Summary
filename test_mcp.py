import subprocess
import json

def main():
    exe_path = r"C:\Users\angelahn_wang\.local\bin\notebooklm-mcp.exe"
    
    # Start the MCP server process
    process = subprocess.Popen(
        [exe_path],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )
    
    # 1. Initialize
    init_req = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "test-client", "version": "1.0.0"}
        }
    }
    process.stdin.write(json.dumps(init_req) + "\n")
    process.stdin.flush()
    
    # Read initialize response
    init_resp = process.stdout.readline()
    print("Initialize Response:")
    print(init_resp)
    
    # Send initialized notification
    init_notif = {
        "jsonrpc": "2.0",
        "method": "notifications/initialized"
    }
    process.stdin.write(json.dumps(init_notif) + "\n")
    process.stdin.flush()
    
    # 2. List tools
    list_tools_req = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/list"
    }
    process.stdin.write(json.dumps(list_tools_req) + "\n")
    process.stdin.flush()
    
    # Read list tools response
    list_tools_resp = process.stdout.readline()
    print("\nTools List Response:")
    print(list_tools_resp)
    
    # Close process
    process.terminate()

if __name__ == "__main__":
    main()
