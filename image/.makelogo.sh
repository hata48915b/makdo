#!/bin/sh

convert -stroke white -strokewidth 10 \
	\( -size 256x256 xc:black -gravity center -pointsize 128 -fill white -annotate +0+0 'M' \) \
	\( -size 256x256 xc:blue  -gravity center -pointsize 128 -fill white -annotate +0+0 'W' \) \
	+append \
	-stroke white -strokewidth 5 \
	-pointsize 96 -fill white -annotate +128-30 '→' \
	-pointsize 96 -fill white -annotate +128+10 '←' \
	-depth 2 \
	md8docx.xpm

convert -stroke white -strokewidth 10 \
	\( -size 256x256 xc:black -gravity center -pointsize 128 -fill white -annotate +0+0 'M' \) \
	\( -size 256x256 xc:blue  -gravity center -pointsize 128 -fill white -annotate +0+0 'W' \) \
	+append \
	-stroke white -strokewidth 5 \
	-pointsize 96 -fill white -annotate +128-30 '→' \
	-pointsize 96 -fill white -annotate +128+10 '←' \
	-resize 256x256! \
	-depth 2 \
	md4docx.xpm

convert -depth 2 md8docx.xpm md8docx.png
convert -depth 2 md4docx.xpm md4docx.png

rm md8docx.xpm
rm md4docx.xpm

