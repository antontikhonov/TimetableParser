enum class ClassNumber(val time: String) {

    FIRST("1. 08:30-10:00"),
    SECOND("2. 10:10-11:40"),
    THIRD("3. 12:20-13:50"),
    FOURTH("4. 14:00-15:30"),
    FIFTH("5. 15:40-17:10"),
    SIXTH("6. 17:20-18:50"),
    SEVENTH("7. 19:00-20:30");

    companion object {
        fun getNumber(number: Int): ClassNumber {
            return when (number) {
                1 -> FIRST
                2 -> SECOND
                3 -> THIRD
                4 -> FOURTH
                5 -> FIFTH
                6 -> SIXTH
                7 -> SEVENTH
                else -> throw Exception("this number not found")
            }
        }
    }
}