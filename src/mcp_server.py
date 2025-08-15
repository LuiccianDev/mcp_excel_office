from mcp_excel.main import run_server


if __name__ == "__main__":
    mcp = run_server()
    mcp.run(transport="stdio")
