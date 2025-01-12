% Name:         .makdo.mp
% Version:      v01
% Time-stamp:   <2024.09.23-12:50:51-JST>

prologues := 3;

beginfig(-1);

  pickup pencircle scaled 172;

  color  bk;  bk = 0.125red + 0.125green + 0.125blue;
  color mbk; mbk = 0.125red + 0.125green + 0.000blue; % 202020200000 / 202000
  color  bl;  bl = 0.125red + 0.375green + 0.750blue;
  color mbl; mbl = 0.125red + 0.375green + 0.625blue; % 202060609F9F / 20609F

  path canvas_lft;
  canvas_lft = (   0,   0)--(   0,2048)--(1024,2048)--(1024,   0)--cycle;
  fill canvas_lft withcolor mbk;

  path canvas_rgt;
  canvas_rgt = (2048,   0)--(2048,2048)--(1024,2048)--(1024,   0)--cycle;
  fill canvas_rgt withcolor mbl;

  path circle_lft;
  circle_lft = (1024,   0)..(  0,1024)..(1024,2048)--cycle;
  fill circle_lft withcolor bk;

  path circle_rgt;
  circle_rgt = (1024,   0)..(2048,1024)..(1024,2048)--cycle;
  fill circle_rgt withcolor bl;

  path halfcircle_lft;
  halfcircle_lft = (1024,   0)..( 512, 512)..(1024,1024)--cycle;
  fill halfcircle_lft withcolor bl;

  path halfcircle_rgt;
  halfcircle_rgt = (1024,1024)..(1536,1536)..(1024,2048)--cycle;
  fill halfcircle_rgt withcolor bk;

  %path quartercircle_up;
  %quartercircle_up = (1024-128,1536)..(1024+128,1536)..cycle;
  %fill quartercircle_up withcolor bl;

  %path quartercircle_dn;
  %quartercircle_dn = (1024-128, 512)..(1024+128, 512)..cycle;
  %fill quartercircle_dn withcolor bk;

  path arrow_pole_lft;
  arrow_pole_lft = ( 512,1024){0,1}..{1,0}(1024,1536);
  draw arrow_pole_lft withcolor white;

  path arrow_pole_rgt;
  arrow_pole_rgt = (1536,1024){0,-1}..{-1,0}(1024, 512);
  draw arrow_pole_rgt withcolor white;

  % 512 * 1 / 3=170.66666666666666666666666666666666
  % 512 * 2 / 3=341.33333333333333333333333333333333
  % 512 / sqrt(3)=295.60333782508839142868417561700262

  x = 171;
  y = 341;
  l = 295;

   path arrow_head_lft;
  arrow_head_lft = (1024-x,1536-l)--(1024-x,1536+l)--(1024+y,1536)--cycle;
  fill arrow_head_lft withcolor white;

  path arrow_head_rgt;
  arrow_head_rgt = (1024+x, 512-l)--(1024+x, 512+l)--(1024-y, 512)--cycle;
  fill arrow_head_rgt withcolor white;

endfig;
end.
