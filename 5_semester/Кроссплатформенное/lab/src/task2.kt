import java.util.stream.Stream
import java.util.stream.Stream.*

fun main() {

    println("Сумма чисел: ${addUp(listOf(1,2,3,4,5).stream())}")

    val artists = Stream.of(
        Artist("The Beatles", "Liverpool"),
        Artist("Queen", "London"),
        Artist("Nirvana", "Seattle"),
        Artist("Madonna", "Michigan"),
        Artist("Elvis Presley", "Mississippi"),
        Artist("Bob Marley", "Jamaica")
    )

    val artistsInfo = getArtistsInfo(artists)
    println("\nИнформация об артистах:")
    artistsInfo.forEach { println(it) }


    val albums = Stream.of(
        Album("Album 1", "Artist A", listOf(
            Track("Song 1", 180),
            Track("Song 2", 240)
        )),
        Album("Album 2", "Artist B", listOf(
            Track("Track 1", 200),
            Track("Track 2", 220),
            Track("Track 3", 190)
        )),
        Album("Album 3", "Artist C", listOf(
            Track("Intro", 60),
            Track("Verse", 180),
            Track("Chorus", 200),
            Track("Outro", 90)
        )),
        Album("Single", "Artist D", listOf(
            Track("Main Track", 300)
        )),
        Album("Greatest Hits", "Artist E", listOf(
            Track("Hit 1", 240),
            Track("Hit 2", 210),
            Track("Hit 3", 195),
            Track("Hit 4", 220),
            Track("Hit 5", 230)
        ))
    )

    val shortAlbums = getAlbumsWithMaxThreeTracks(albums)
    println("\nАльбомы с не более чем 3 треками:")
    shortAlbums.forEach { album ->
        println("${album.title} - ${album.artist} (${album.tracks.size} треков)")
    }
}

fun addUp(numbers: Stream<Int>): Int {
    return numbers.mapToInt { it }.sum()
}

data class Artist(
    val name: String,
    val origin: String
)

fun getArtistsInfo(artists: Stream<Artist>): List<String> {
    return artists
        .map { artist -> "${artist.name} - ${artist.origin}" }
        .toList()
}

data class Track(
    val title: String,
    val duration: Int
)

data class Album(
    val title: String,
    val artist: String,
    val tracks: List<Track>
)

fun getAlbumsWithMaxThreeTracks(albums: Stream<Album>): List<Album> {
    return albums
        .filter { album -> album.tracks.size <= 3 }
        .toList()
}