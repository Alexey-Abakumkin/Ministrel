# Ministrel

The scripts creates pptx presentation from text file with lyrics

# Requirements

[python-pptx](https://python-pptx.readthedocs.org/en/latest/user/install.html#install)

## Usage

```
usage: ministrel.py [-h] [-m SLIDE] [-c {white,black}] INPUT OUTPUT

positional arguments:
  INPUT                 text file name contains lyrics of the song
  OUTPUT                output presentation destination

optional arguments:
  -h, --help            show this help message and exit
  -m SLIDE, --master SLIDE
                        destination to master pptx slide (default:
                        "master.pptx")
  -c {white,black}, --font-color {white,black}
                        color of the presentation font (default: "white")
```                        

## Example

Input file ([link](https://raw.githubusercontent.com/Alexey-Abakumkin/Ministrel/master/example/song.txt)):

```
So you run and you run to catch up with the sun but it's sinking
Racing around to come up behind you again.
The sun is the same in a relative way but you're older,
Shorter of breath and one day closer to death.

Every year is getting shorter; never seem to find the time.
Plans that either come to naught or half a page of scribbled lines
Hanging on in quiet desperation is the English way
The time is gone, the song is over,
Thought I'd something more to say.

Home
Home again
I like to be here
When I can
```

Output: [presentation](https://github.com/Alexey-Abakumkin/Ministrel/blob/master/example/time.pptx?raw=true)
