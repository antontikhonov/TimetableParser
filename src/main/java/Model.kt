data class GroupTimetable(
    val number_of_group: String,
    val days: List<TimetableDay>
)

data class TimetableDay(
    val day: String,
    val classes: List<PairKlass>
)

data class PairKlass(
    val time: String,
    val odd: String?,
    val even: String?
)