nv = 1 # Номер по журналу
ngl = 1 #
ngp = 2

def to_hex_twos_complement(num, bits=32):
    input_num = num
    # Определяем минимальный тип данных для числа
    min_bytes = None

    # Проверяем в обратном порядке (от 8 байт к 1)
    if -9223372036854775808 <= num <= 9223372036854775807:  # 64-bit (8 bytes)
        min_bytes = 8
        if -2147483648 <= num <= 2147483647:  # 32-bit (4 bytes)
            if -32768 <= num <= 32767:  # 16-bit (2 bytes)
                if -128 <= num <= 127:  # 8-bit (1 byte)
                    min_bytes = "1 (byte)"
                else:
                    min_bytes = "2 (word)"
            else:
                min_bytes = "4 (long word)"

    mask = (1 << bits) - 1
    if num < 0:
        num = (1 << bits) + num
    num = num & mask
    hex_digits = bits // 4
    hex_str = format(num, f'0{hex_digits}X')
    print(f"исх.: {input_num} ; 16-чн.: 0x{hex_str} ; нужное кол-во байт: {min_bytes}")

x1 = (-1)**nv*((nv+ngp)*3)
x2 = (-1)**(nv+1)*((nv+ngl)*2)
x3 = (-1)**(nv+ngl)*(nv+ngp+ngl+21)
x4 = (-1)**nv*(nv+ngp+29)**2
x5 = (-1)**(nv+1)*(nv+ngl+23)**2
x6 = (-1)**(nv+ngl)*(nv+ngp+ngl+79)**2
x7 = x4**2
x8 = -x5**2
x9 = (-1)**(nv+ngl)*x6**2

to_hex_twos_complement(x1)
to_hex_twos_complement(x2)
to_hex_twos_complement(x3)
to_hex_twos_complement(x4)
to_hex_twos_complement(x5)
to_hex_twos_complement(x6)
to_hex_twos_complement(x7)
to_hex_twos_complement(x8)
to_hex_twos_complement(x9)
