#!/bin/sh

TMPDIR=tmp
test -e $TMPDIR && exit 1
mkdir $TMPDIR

cp -p .makdo.mp x0.mp
upmpost x0.mp
convert -depth 8 -colors 512 x0.ps x0.xpm

cat x0.xpm \
    | sed 's;#202020200000";None";g' | sed 's;#202000";None";g' \
    | sed 's;#202060609F9F";None";g' | sed 's;#20609F";None";g' \
					 > xB.xpm

# MAKE THE APPLI IMAGE

# convert -quality 100 -resize 2048x2048! -depth 8 -colors 512 xB.xpm makdoB.png; chmod 444 makdoB.png
# convert -quality 100 -resize 1024x1024! -depth 8 -colors 512 xB.xpm makdoA.png; chmod 444 makdoA.png
# convert -quality 100 -resize  512x512!  -depth 8 -colors 512 xB.xpm makdo9.png; chmod 444 makdo9.png
convert -quality 100 -resize  256x256!  -depth 8 -colors 512 xB.xpm makdo8.png; chmod 444 makdo8.png
# convert -quality 100 -resize  128x128!  -depth 8 -colors 512 xB.xpm makdo7.png; chmod 444 makdo7.png
# convert -quality 100 -resize   64x64!   -depth 8 -colors 512 xB.xpm makdo6.png; chmod 444 makdo6.png
# convert -quality 100 -resize   32x32!   -depth 8 -colors 512 xB.xpm makdo5.png; chmod 444 makdo5.png
# convert -quality 100 -resize   16x16!   -depth 8 -colors 512 xB.xpm makdo4.png; chmod 444 makdo4.png
convert -quality 100 -resize  512x256!  -depth 8 -colors 512 xB.xpm makdoL.png; chmod 444 makdoL.png

# MAKE THE SPLASH IMAGE

cat x0.xpm \
    | sed 's;#202020200000";#7FFF7FFF7FFF";g' | sed 's;#202000";#7F7F7F";g' \
    | sed 's;#202060609F9F";#7FFF7FFF7FFF";g' | sed 's;#20609F";#7F7F7F";g' \
					 > xB.xpm

convert -quality 100 \
	'(' -size 512x320 xc:#7F7F7F ')' \
	'(' -resize  256x256! -depth 8 xB.xpm ')' \
	-geometry +128+32 -composite \
	-font IPAゴシック -pointsize 22 -fill '#FFFF00' -stroke '#FFFF00' -strokewidth 0 -annotate +12+25   'MAKDO is starting.' \
	-font IPAゴシック -pointsize 22 -fill '#FFFF00' -stroke '#FFFF00' -strokewidth 0 -annotate +220+313 'Powered by Seiichiro HATA.' \
	makdo_splash.png
chmod 444 makdo_splash.png

# MAKE THE ICONS

cat x0.xpm \
    | sed 's;#202020200000";None";g' | sed 's;#202000";None";g' \
    | sed 's;#202060609F9F";None";g' | sed 's;#20609F";None";g' \
					 > xB.xpm

#convert -quality 100 -resize 2048x2048! -depth 8 -colors 512 xB.xpm $TMPDIR/icoB.png
convert -quality 100 -resize 1024x1024! -depth 8 -colors 512 xB.xpm $TMPDIR/icoA.png
convert -quality 100 -resize  512x512!  -depth 8 -colors 512 xB.xpm $TMPDIR/ico9.png
convert -quality 100 -resize  256x256!  -depth 8 -colors 512 xB.xpm $TMPDIR/ico8.png
convert -quality 100 -resize  128x128!  -depth 8 -colors 512 xB.xpm $TMPDIR/ico7.png
#convert -quality 100 -resize   64x64!   -depth 8 -colors 512 xB.xpm $TMPDIR/ico6.png
convert -quality 100 -resize   32x32!   -depth 8 -colors 512 xB.xpm $TMPDIR/ico5.png
convert -quality 100 -resize   16x16!   -depth 8 -colors 512 xB.xpm $TMPDIR/ico4.png
#convert -quality 100 -resize  512x256!  -depth 8 -colors 512 xB.xpm $TMPDIR/icoL.png

# WINDOWS
convert $TMPDIR/icoA.png -define icon:auto-resize=256,128,64,32,16 makdo_win.ico
chmod 444 makdo_win.ico
# MACOS
png2icns makdo_mac.icns $TMPDIR/ico?.png
chmod 444 makdo_mac.icns

# CLEAN FILES

rm x0.log x0.mp x0.ps x0.xpm xB.xpm

rm $TMPDIR/ico?.png
rmdir $TMPDIR
