import kotlin.math.*

fun main() {
    print("Введите значение x: ")
    val x = readlnOrNull()?.toDoubleOrNull() ?: return
    print("Введите точность: ")
    val epsilon = readlnOrNull()?.toDoubleOrNull() ?: return

    val result = calcEpsilon(x, epsilon)
    println("Значение e в степени $x с точностью $epsilon: $result")
}

fun factorial(n: Int): Long {
    if (n == 0) {
        return 1
    }
    return n * factorial(n - 1)
}

fun task(x: Double, n: Int): Double {
    return x.pow(n) / factorial(n).toDouble()
}

fun calcEpsilon(x: Double, epsilon: Double): Double {
    var result = 1.0
    var n = 1

    while (true) {
        val local = task(x, n)
        if (local < epsilon) {
            break
        }
        result += local
        n++
    }

    return result
}