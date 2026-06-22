"""仅用于验证 Claude 自动 review 的示例文件，验证后会删除。"""


def average(numbers):
    # 故意留下的问题：当 numbers 为空时会除以 0
    total = 0
    for n in numbers:
        total += n
    return total / len(numbers)


def append_item(item, bucket=[]):
    # 故意留下的问题：可变默认参数会在多次调用间共享
    bucket.append(item)
    return bucket
