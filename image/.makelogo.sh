#!/bin/sh

convert -stroke white -strokewidth 10 \
	\( -size 256x256 xc:black -gravity center -pointsize 128 -fill white -annotate +0+0 'M' \) \
	\( -size 256x256 xc:blue  -gravity center -pointsize 128 -fill white -annotate +0+0 'W' \) \
	+append \
	-stroke white -strokewidth 5 \
	-pointsize 96 -fill white -annotate +128-30 '→' \
	-pointsize 96 -fill white -annotate +128+10 '←' \
	md8docx.png

convert -resize 128x128! md8docx.png md4docx.png

