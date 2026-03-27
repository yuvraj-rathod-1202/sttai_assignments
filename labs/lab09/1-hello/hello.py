import sys, platform
print(f"Python: {sys.version}")
print(f"OS: {platform.system()} {platform.release()}")
print(f"Architecture: {platform.machine()}")
print("Hello from inside Docker! 🐳")