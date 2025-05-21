import time

def count_with_interval():
    for i in range(1, 101):
        print(f"Count: {i}")
        time.sleep(10)  # Wait for 10 seconds

if __name__ == "__main__":
    count_with_interval()