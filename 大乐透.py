from random import sample


def choosenum(redmax=33, redn=6, bluemax=16, bluen=1):
    assert redn <= redmax, '传递参数错误：redn>redmax'
    assert bluen <= bluemax, '传递参数错误：bluen>bluemax'
    redseq = sample(list(range(1, redmax + 1)), redn)
    blueseq = sample(list(range(1, bluemax + 1)), bluen)
    return (redseq, blueseq)


def main():
    while True:
        if input('输入“end”结束，输入其他继续:') == 'end':
            break
        r, b = choosenum()
        print('双色球号码\n红球：{},蓝球：{}'.format(r, b))
        r, b = choosenum(35, 5, 12, 2)
        print('大乐透号码\n红球：{},蓝球：{}'.format(r, b))


if __name__ == '__main__':
    main()