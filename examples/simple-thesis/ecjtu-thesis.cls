\NeedsTeXFormat{LaTeX2e}
\ProvidesClass{ecjtu-thesis}[2026/10/04 ECJTU software engineering bachelor thesis]

\newif\ifecjtuguide
\ecjtuguidetrue
\DeclareOption{final}{\ecjtuguidefalse}
\DeclareOption*{\PassOptionsToClass{\CurrentOption}{ctexbook}}
\ProcessOptions\relax
\LoadClass[UTF8,zihao=-4,a4paper,twoside,openany,fontset=none]{ctexbook}
\RequirePackage{iftex}
\RequireXeTeX
\RequirePackage{geometry}
% 依据学校规范：上下左 25 mm，右 20 mm，奇偶页保持相同版心。
\geometry{paperwidth=595.32bp,paperheight=841.92bp,left=70.944bp,right=56.676bp,top=70.9bp,bottom=70.9bp,
  asymmetric,headheight=19.5bp,headsep=8.85bp,footskip=7.5mm}
\RequirePackage{fontspec}
% 必须使用指定字体；不以替代字体悄悄改变字形与分页。
\setmainfont{Times New Roman}
\setsansfont{Times New Roman}
\setmonofont{Times New Roman}
\setCJKmainfont[AutoFakeBold=2.84647,ItalicFont=KaiTi]{SimSun}
\setCJKsansfont[AutoFakeBold=2.84647]{SimHei}
\setCJKmonofont{SimSun}
\setCJKfamilyfont{zhsong}[AutoFakeBold=2.84647]{SimSun}
\setCJKfamilyfont{zhhei}[AutoFakeBold=2.84647]{SimHei}
% Word 原稿使用“楷体_GB2312”，系统的 KaiTi 字形和字宽并不相同。
% 从用户提供的 Word PDF 提取同一嵌入字体，保证诚信声明换行与原稿一致。
\setCJKfamilyfont{zhkai}[Path=assets/,AutoFakeBold=2.84647]{KaiTi_GB2312.ttf}
\providecommand{\songti}{\CJKfamily{zhsong}}
\providecommand{\heiti}{\CJKfamily{zhhei}}
\providecommand{\kaishu}{\CJKfamily{zhkai}}
\newCJKfontfamily\ecjtuxing[Path=assets/,FakeBold=2.84647,FakeSlant=0.3333]{STXINGKA.TTF}
\newCJKfontfamily\ecjtucoverhintfont[FakeBold=2.84647,FakeSlant=0.3333]{SimSun}
\newfontfamily\ecjtucoverhintlatin[FakeBold=2.84647,FakeSlant=0.3333]{SimSun}
\newfontfamily\ecjtutitlelatin[AutoFakeBold=2.84647]{SimHei}
\newCJKfontfamily\ecjtuguidanceemphasis[FakeBold=2.84647,FakeSlant=0.3333]{SimSun}
\newfontfamily\ecjtuguidanceemlatin{Times New Roman Bold Italic}
\newcommand{\guidanceem}[1]{{\ecjtuguidanceemphasis\ecjtuguidanceemlatin #1}}
\RequirePackage{amsmath,amssymb}
\RequirePackage{graphicx,xcolor,tikz}
\usetikzlibrary{arrows.meta,positioning,shapes.geometric,calc}
\RequirePackage{array,tabularx,longtable,booktabs}
\RequirePackage{caption}
\RequirePackage{float,placeins}
\RequirePackage{fancyhdr}
\RequirePackage{titletoc}
\RequirePackage{enumitem}
\RequirePackage{gbt7714}
\RequirePackage{xurl}
\RequirePackage[hidelinks,unicode,bookmarksnumbered]{hyperref}

% Word 导出的内容页显式绘制白色纸面，并按纸张边界裁切。
\AddToHook{shipout/background}{%
  \ifnum\ReadonlyShipoutCounter=2\else\ifnum\ReadonlyShipoutCounter=4\else
    \special{pdf:put @thispage << /Group << /Type /Group /S /Transparency /CS /DeviceRGB >> >>}%
    \begin{tikzpicture}[remember picture,overlay]
      \special{pdf:code q 0 -841.919938965 595.32 841.92 re W* n
        1 g 0 -842.05 595.45 842.05 re f* Q}%
    \end{tikzpicture}%
  \fi\fi}

% Word 小四正文单倍基准 15.6 bp，乘 1.25 后基线距离为 19.5 bp。
\renewcommand\normalsize{\@setfontsize\normalsize{12bp}{19.5bp}%
  \abovedisplayskip=9.75bp plus 2bp minus 2bp
  \belowdisplayskip=9.75bp plus 2bp minus 2bp
  \abovedisplayshortskip=0bp plus 2bp
  \belowdisplayshortskip=6bp plus 2bp minus 2bp}
\linespread{1}
\AtBeginDocument{\normalsize\songti\color{black}}
\setlength{\parindent}{2em}
\setlength{\parskip}{0bp}
% Word 的行框随字号变化：12 / 10.56 bp 正文为 19.5 bp，
% 14.04 / 15.96 bp 标题为 31.2 bp。用实测的基线上下沿连接行框，
% 避免 LaTeX 相邻标题自动吞掉段前间距而逐节累积偏移。
\newif\ifecjtuwordgrid
\newcommand{\ecjtuwordstrut}[2]{\rule[-#2bp]{0bp}{\dimexpr#1bp+#2bp\relax}}
\newcommand{\ecjtutextstrut}{%
  \ifdim\f@size pt<11bp\ecjtuwordstrut{13.6}{5.9}%
  \else\ecjtuwordstrut{14.08}{5.42}\fi}
\AddToHook{para/begin}{\ifecjtuwordgrid\ecjtutextstrut\fi}
\AddToHook{para/end}{\ifecjtuwordgrid\ecjtutextstrut\fi}
\setlength{\lineskip}{0bp}
\newcommand{\ecjtuheadingpreskip}{%
  \ifdim\prevdepth>-999pt
    \dimexpr\ifecjtusmallhints5.9bp\else5.42bp\fi-\prevdepth\relax
  \else0bp\fi}
\newcommand{\ecjtusectionlatin}{\ifnum\value{chapter}<4\rmfamily\else\ecjtutitlelatin\fi}
\newcommand{\ecjtuchapterlabel}[1]{{\XeTeXinterchartokenstate=0
  第\hspace{4.08bp}\arabic{chapter}\hspace{\ifnum\value{chapter}=4 8.16bp\else4.14bp\fi}%
  {\ifnum\value{chapter}=1\mdseries\fi\ifnum\value{chapter}=3\mdseries\fi 章}}}
\DeclareRobustCommand{\thesistitlespace}[1][8.16bp]{\ifdim\f@size pt>13bp\hspace{#1}\else\hspace{6bp}\fi}
\newif\ifecjtusmallhints
\newcommand{\smalltemplatehints}{\ecjtusmallhintstrue}
\newcommand{\normaltemplatehints}{\ecjtusmallhintsfalse}
\newif\ifecjtusimsunhints
\newcommand{\simsuntemplatehints}{\ecjtusimsunhintstrue}
\newcommand{\normaltemplatehintfont}{\ecjtusimsunhintsfalse}
\newif\ifecjtuindenthints
\newcommand{\indenttemplatehints}{\ecjtuindenthintstrue}
\newcommand{\normaltemplatehintindent}{\ecjtuindenthintsfalse}
\setlength{\emergencystretch}{1em}
\raggedbottom
\clubpenalty=10000
\widowpenalty=10000
\displaywidowpenalty=10000
\ctexset{
  chapter={name={第,章},number=\arabic{chapter},
    format=\centering\heiti\ecjtutitlelatin\bfseries\fontsize{15.96bp}{31.2bp}\selectfont\ecjtuwordstrut{21.4}{9.8},
    nameformat=\ecjtuchapterlabel,titleformat={},aftername=\hspace{8.16bp},
    beforeskip=0bp,afterskip=0bp,fixskip=true,pagestyle=thesis},
  section={format=\ecjtusectionlatin\heiti\bfseries\fontsize{14.04bp}{31.2bp}\selectfont\ecjtuwordstrut{20.68}{10.52},
    aftername=\hspace{0.5em},beforeskip=\ecjtuheadingpreskip,afterskip=0bp,fixskip=true},
  subsection={format=\ecjtusectionlatin\heiti\mdseries\fontsize{12bp}{19.5bp}\selectfont
    \ifnum\value{chapter}>3\xeCJKsetup{CJKecglue=\hskip3bp}\fi\ecjtuwordstrut{14.08}{5.42},
    aftername=\hspace{0.5em},beforeskip=\ecjtuheadingpreskip,afterskip=0bp,fixskip=true}}
\setcounter{secnumdepth}{2}
\setcounter{tocdepth}{2}

\newcommand{\thesissetup}[1]{\setkeys{ecjtu}{#1}}
\define@key{ecjtu}{title}{\def\ecjtuTitle{#1}}
\define@key{ecjtu}{title-en}{\def\ecjtuTitleEn{#1}}
\define@key{ecjtu}{author}{\def\ecjtuAuthor{#1}}
\define@key{ecjtu}{student-id}{\def\ecjtuStudentID{#1}}
\define@key{ecjtu}{school}{\def\ecjtuSchool{#1}}
\define@key{ecjtu}{major}{\def\ecjtuMajor{#1}}
\define@key{ecjtu}{class}{\def\ecjtuClass{#1}}
\define@key{ecjtu}{supervisor}{\def\ecjtuSupervisor{#1}}
\define@key{ecjtu}{date}{\def\ecjtuDate{#1}}
\define@key{ecjtu}{sign-date}{\def\ecjtuSignDate{#1}}
\thesissetup{title={毕业论文题目},title-en={Thesis Title},author={学生姓名},
  student-id={学号},school={信息与软件工程学院},major={软件工程},
  class={班级},supervisor={指导教师},date={2026年\quad 月},
  sign-date={2026 年\quad 月\quad 日}}

\newcommand{\ecjtuprintpagenumber}{%
  \ifecjtuguide\ifnum\value{page}>9
    {\addfontfeatures{LetterSpace=0.666667}\makebox[9bp][l]{\thepage}}%
  \else\thepage\fi\else\thepage\fi}

\fancypagestyle{thesis}{%
  \fancyhf{}
  \fancyhfinit{\ecjtuwordgridfalse}
  \fancyhead[C]{\ecjtuwordgridfalse
    \def\ecjtuheadercenter{304.73bp}%
    \def\ecjtufootercenter{304.73bp}%
    \ifecjtuguide\ifnum\value{page}>9\def\ecjtufootercenter{304.81bp}\fi\fi
    \ifodd\value{page}\else\ifecjtuguide\def\ecjtuheadercenter{302.183bp}\fi\fi
    \begin{tikzpicture}[remember picture,overlay]
    \node[anchor=base,inner sep=0,outer sep=0,
      font=\songti\mdseries\fontsize{9bp}{12bp}\selectfont]
      at ([xshift=\ecjtuheadercenter,yshift=-51.72bp]current page.north west)
      {\ifodd\value{page}华东交通大学毕业设计\else
        \ifecjtuguide\color{red}学生姓名：毕业论文题目（请将这里修改为自己的姓名和题目，并将括号内容删除。）\else
          \ecjtuAuthor：\ecjtuTitle\fi\fi};
    \fill ([xshift=69.504bp,yshift=-55.2bp]current page.north west)
      rectangle ([xshift=540.124bp,yshift=-55.92bp]current page.north west);
    \node[anchor=base,inner sep=0,outer sep=0,
      font=\rmfamily\mdseries\fontsize{9bp}{12bp}\selectfont]
      at ([xshift=\ecjtufootercenter,yshift=-790.44bp]current page.north west) {\ecjtuprintpagenumber};
  \end{tikzpicture}}
  \renewcommand{\headrulewidth}{0bp}
  \renewcommand{\footrulewidth}{0bp}}
\fancypagestyle{plain}{\pagestyle{thesis}}

% 目录按学院 Word 批注：三号加粗标题；小四内容，逐级缩进一字。
% ctex 已把“第…章 / 附录…”写入 numberline，直接使用现成编号。
\newif\ifecjtutocprevchapter
\def\ecjtutocchapnum{0}
\def\ecjtutocdetect#1#2#3\@nil{\ifx#1第\gdef\ecjtutocchapnum{#2}\else\gdef\ecjtutocchapnum{0}\fi}
\newcommand{\ecjtutocsetchapter}{\expandafter\ecjtutocdetect\thecontentslabel\@nil}
\newcommand{\ecjtutocnumberfont}{\ifnum\ecjtutocchapnum<4\rmfamily\else\ecjtutitlelatin\fi}
\newcommand{\ecjtutocpagenumberwidth}{%
  \settowidth{\@tempdima}{\thecontentspage}%
  \edef\ecjtutocpagewidth{\the\@tempdima}%
  \contentsmargin{\ecjtutocpagewidth}}
\newcommand{\ecjtutocsectionextra}{%
  \def\ecjtutocspaceextra{0bp}%
  \ifnum\ecjtutocchapnum=2
    \edef\ecjtutoccurrentlabel{\thecontentslabel}\def\ecjtutoclastlabel{2.4}%
    \ifx\ecjtutoccurrentlabel\ecjtutoclastlabel\else\def\ecjtutocspaceextra{6bp}\fi
  \fi
  \addtolength{\leftskip}{\ecjtutocspaceextra}}
\newcommand{\ecjtutocdots}[1][-1.86bp]{%
  \hspace{#1}\titlerule*[3bp]{\makebox[3bp][l]{\hspace{1.14bp}.}}}
\newcommand{\ecjtutocchapterdots}{%
  \ifnum\ecjtutocchapnum<2\ecjtutitlelatin
    \hspace{-3.58bp}\titlerule*[6.12bp]{\makebox[6.12bp][l]{\hspace{2.54bp}.}}%
  \else\rmfamily\ecjtutocdots\fi\contentspage}
\titlecontents{chapter}[42.12bp]{\heiti\ecjtutitlelatin\bfseries\fontsize{12bp}{15.6bp}\selectfont
    \ecjtutocpagenumberwidth
    \ifecjtutocprevchapter\else\addvspace{6bp}\fi\global\ecjtutocprevchaptertrue}
  {\ecjtutocsetchapter\contentslabel[\xeCJKsetup{CJKecglue=\hskip3.06bp}%
    \ifnum\ecjtutocchapnum=4{\XeTeXinterchartokenstate=0 第\hspace{3.06bp}4\hspace{6.06bp}章}\else\thecontentslabel\fi]{42.12bp}%
    \ifnum\ecjtutocchapnum=4\hspace{3.12bp}\fi\ifnum\ecjtutocchapnum=0\hspace{-2.88bp}\fi}
  {\gdef\ecjtutocchapnum{0}\hspace*{-42.12bp}}
  {\ecjtutocchapterdots}[\addvspace{6bp}]
\titlecontents{section}[31.56bp]
  {\ifnum\ecjtutocchapnum>3\addtolength{\leftskip}{3bp}\fi
    \ecjtutocsectionextra
    \global\ecjtutocprevchapterfalse\songti\rmfamily\mdseries\fontsize{12bp}{15.6bp}\selectfont\ecjtutocpagenumberwidth}
  {\contentslabel[\ecjtutocnumberfont\thecontentslabel]{\dimexpr\ifnum\ecjtutocchapnum<4 21bp\else24bp\fi+\ecjtutocspaceextra\relax}}{}
  {\ecjtutocdots[0bp]\contentspage}
\titlecontents{subsection}[51bp]
  {\ifnum\ecjtutocchapnum>3\addtolength{\leftskip}{6bp}\fi
    \global\ecjtutocprevchapterfalse\songti\rmfamily\mdseries\fontsize{12bp}{15.6bp}\selectfont
    \ifnum\ecjtutocchapnum>3\ecjtutitlelatin\xeCJKsetup{CJKecglue=\hskip3bp}\fi\ecjtutocpagenumberwidth}
  {\contentslabel[\ecjtutocnumberfont\thecontentslabel]{\ifnum\ecjtutocchapnum<4 30bp\else36bp\fi}}{}
  {\rmfamily\ecjtutocdots\contentspage}

\DeclareCaptionFont{ecjtu}{\heiti\mdseries\fontsize{10.56bp}{19.5bp}\selectfont\xeCJKsetup{CJKglue=\hskip-0.06bp}}
\DeclareCaptionLabelFormat{ecjtu}{#1\hspace{5.28bp}#2}
\DeclareCaptionLabelSeparator{ecjtuspace}{\hspace{5.28bp}}
\captionsetup{font=ecjtu,labelfont=ecjtu,textfont=ecjtu,
  labelformat=ecjtu,labelsep=ecjtuspace,justification=centering,singlelinecheck=true,skip=6bp}
\captionsetup[table]{position=top}
\renewcommand{\thefigure}{\thechapter-\arabic{figure}}
\renewcommand{\thetable}{\thechapter-\arabic{table}}
\renewcommand{\theequation}{\thechapter-\arabic{equation}}
\setlength{\textfloatsep}{12bp}
\setlength{\intextsep}{12bp}
\newcommand{\tablefont}{\songti\mdseries\fontsize{10.56bp}{19.5bp}\selectfont\xeCJKsetup{CJKglue=\hskip-0.06bp}}
\newcolumntype{Y}{>{\raggedright\arraybackslash}X}
\setlist{nosep,leftmargin=2em,labelsep=0.5em,topsep=0bp,partopsep=0bp}
% 原 Word 模板把写作提示直接印在稿面上：说明为红色，样例为蓝色。
% [final] 选项仍可一次性隐藏这些提示。
\newcommand{\templatehint}[1]{\ifecjtuguide\par\begingroup
  \color{red}\frenchspacing%
  \ifecjtusmallhints\else\ifecjtusimsunhints\else\xeCJKsetup{PunctStyle=plain}\fi\fi
  \ifecjtusimsunhints\fontspec{SimSun}\XeTeXcharclass"00B7=0\relax
    \xeCJKsetup{PunctStyle=quanjiao}%
    \hyphenpenalty=10000\exhyphenpenalty=10000\fi
  \ifecjtusmallhints\fontsize{10.56bp}{19.5bp}\selectfont\xeCJKsetup{CJKglue=\hskip-0.06bp}\fi
  \noindent\ifecjtuindenthints\hspace*{12bp}\fi #1\par\endgroup\fi}
\newcommand{\templateheading}[1]{\templatehint{#1}\nobreak}
\newcommand{\templateblank}{}
\newcommand{\keywords}[1]{\par\vspace{39bp}\noindent{\ifecjtuguide\color{blue}\fi\songti 关键词：#1}\par}
\newcommand{\keywordsen}[1]{\par\vspace{58.5bp}\noindent{{\color{black}\fontspec{SimHei} Keywords}{\color{black}\rmfamily: }\ifecjtuguide\color{blue}\fi\rmfamily #1}\par}

\newcommand{\makecover}{%
  \clearpage\thispagestyle{empty}
  \begin{tikzpicture}[remember picture,overlay]
    % 原 Word 图片 8：相对正文左移 105 pt、上移 85.8 pt，598.5 × 858 pt。
    \node[anchor=north west,inner sep=0,outer sep=0]
      at ([xshift=-34.1bp,yshift=14.9bp]current page.north west)
      {\includegraphics[width=598.5bp,height=858bp]{assets/cover-background.png}};
    \begin{scope}[shift={(current page.north west)},x=1mm,y=1mm]
      % 指定 Word 模板的默认稿面保留空白填写线；正式稿 [final] 才填入元数据。
      \ifecjtuguide\else
        \node[anchor=south,align=center,text width=113mm,inner sep=0,
          font=\songti\fontsize{14bp}{20bp}\selectfont] at (122.18,-131.30) {\ecjtuTitle};
        \node[anchor=south,text width=76mm,align=center,inner sep=0,
          font=\songti\fontsize{14bp}{20bp}\selectfont] at (116.65,-151.68) {\ecjtuSchool};
        \node[anchor=south,text width=76mm,align=center,inner sep=0,
          font=\songti\fontsize{14bp}{20bp}\selectfont] at (116.65,-164.62) {\ecjtuMajor\ecjtuClass};
        \node[anchor=south,text width=76mm,align=center,inner sep=0,
          font=\songti\fontsize{14bp}{20bp}\selectfont] at (116.65,-177.56) {\ecjtuAuthor};
        \node[anchor=south,text width=76mm,align=center,inner sep=0,
          font=\songti\fontsize{14bp}{20bp}\selectfont] at (116.65,-191.22) {\ecjtuSupervisor};
        \node[anchor=north,align=center,font=\songti\fontsize{12bp}{19.5bp}\selectfont]
          at (115,-213) {学号：\ecjtuStudentID\\完成日期：\ecjtuDate};
      \fi
      \ifecjtuguide
        \node[anchor=north west,inner sep=0,text width=158mm,
          font=\ecjtucoverhintfont\ecjtucoverhintlatin\fontsize{14.04bp}{31.2bp}\selectfont,align=left]
          at (29.8,-13.585) {\color{red}本论文模版章节设置和内容安排仅作参考，作者可以根据系统自身特点以\\
            及本人的专长做适当调整！\\
            \color{blue}1、蓝色表示样例信息\\
            2、红色表示说明文字\\
            3、彩色文字最后都要删除\\
            4、毕业设计论文中所有文字都用\hspace{7.02bp}黑色};
        \draw[line width=0.75bp] (77.25bp,-25.60bp) rectangle (540bp,-211.35bp);
        \node[anchor=base west,inner sep=0,outer sep=0,text=red]
          at (82.104bp,-598.42bp) {\begin{minipage}[t]{142mm}
            \ecjtuxing\fontsize{21.96bp}{31.2bp}\selectfont\setlength{\parindent}{0bp}%
            \xeCJKsetup{PunctStyle=plain,CJKglue=\hskip0.0568bp}%
            本彩页封面仅为电子稿样本，文本装订时\\[-0.12bp]
            请以班为单位到教材科领取纸质封面及资\\
            料袋后再装订！\end{minipage}};
        \draw[line width=0.75bp] (74.5bp,-573.5bp) rectangle (489.25bp,-682.70bp);
      \fi
    \end{scope}
  \end{tikzpicture}\null\clearpage}

\newcommand{\makedeclaration}{%
  \thispagestyle{empty}
  \begin{tikzpicture}[remember picture,overlay]
    \begin{scope}[shift={(current page.north west)},x=1mm,y=1mm]
      \draw[line width=0.48bp,black] (119.30bp,-71.16bp) rectangle (490.42bp,-695.74bp);
      \node[anchor=base west,inner sep=0,outer sep=0,
        font=\heiti\bfseries\fontsize{15.96bp}{31.2bp}\selectfont\xeCJKsetup{CJKglue=\hskip0.11172bp}]
        at (208.37bp,-155.18bp) {毕业设计（论文）诚信声明};
      \node[anchor=base west,inner sep=0,outer sep=0] at (124.58bp,-217.58bp) {\begin{minipage}[t]{127mm}
        \kaishu\bfseries\fontsize{15.96bp}{31.2bp}\selectfont\xeCJKsetup{PunctStyle=plain,CJKglue=\hskip0.11172bp plus0.08\baselineskip}\setlength{\parindent}{32.16bp}
        本人郑重声明：所呈交的毕业设计（论文）是我个人在导师指导下进行的研究工作及取得的研究成果。就我所知，除了文中特别加以标注和致谢的地方外，论文中不包含其他人已经发表和撰写的研究成果，也不包含为获得华东交通大学或其他教育机构的学位或证书所使用过的材料。

        如在文中涉及抄袭或剽窃行为，本人愿承担由此而造成的一切后果及责任。
      \end{minipage}};
      \node[anchor=base west,inner sep=0,outer sep=0,font=\songti\fontsize{15.96bp}{31.2bp}\selectfont]
        at (236.57bp,-560.86bp) {本人签名};
      \node[anchor=base west,inner sep=0,outer sep=0,font=\songti\fontsize{15.96bp}{31.2bp}\selectfont]
        at (236.45bp,-623.26bp) {导师签名};
      \node[anchor=base west,inner sep=0,outer sep=0,font=\songti\bfseries\fontsize{15bp}{24bp}\selectfont]
        at (320.35bp,-685.30bp) {\ecjtuSignDate};
    \end{scope}
  \end{tikzpicture}\null\clearpage}

\newenvironment{cnabstract}{\clearpage\newgeometry{left=70.944bp,right=56.676bp,top=49.774bp,bottom=70.9bp}\thispagestyle{empty}
  {\centering\ifecjtuguide\color{blue}\fi\heiti\fontsize{18bp}{25bp}\selectfont\ecjtuTitle\par}
  \vspace{14.82bp}{\centering\color{black}\heiti\ecjtutitlelatin\fontsize{15.96bp}{26bp}\selectfont\XeTeXinterchartokenstate=0\hspace*{-2.148bp}摘\hspace{15.98bp}要\par}
  \vspace{12.18bp}\normalsize\songti\ifecjtuguide\color{blue}\fi}{\par\color{black}\restoregeometry}
\newenvironment{enabstract}{\clearpage\newgeometry{left=20mm,right=25mm,top=18.514mm,bottom=25mm}\thispagestyle{empty}
  {\centering\ifecjtuguide\color{blue}\fi\rmfamily\bfseries\fontsize{18bp}{31.2bp}\selectfont\ecjtuTitleEn\par}
  \vspace{20.1bp}{\centering\color{black}\rmfamily\bfseries\fontsize{15.96bp}{26bp}\selectfont \ifecjtuguide Abstract\else ABSTRACT\fi\par}
  % Times New Roman 的 f 等字母笔画可伸出字宽，预留边承使墨迹留在版心。
  \vspace{12.2bp}\normalsize\rmfamily\frenchspacing\ifecjtuguide\color{blue}\fi\setlength{\leftskip}{0.078bp}\setlength{\rightskip}{0bp}\hyphenpenalty=10000\exhyphenpenalty=10000}{\par\color{black}\restoregeometry}
\newcommand{\thesistableofcontents}{%
  \clearpage\pagestyle{empty}\vspace*{-38.895bp}\begingroup\xeCJKsetup{PunctStyle=plain}
  {\centering\heiti\ecjtutitlelatin\bfseries\fontsize{15.96bp}{26bp}\selectfont\XeTeXinterchartokenstate=0\hspace*{-0.108bp}目\hspace{16.34bp}录\par}
  \vspace{9.92bp}\ifecjtuguide{\color{red}\songti\bfseries\fontsize{10.56bp}{19.5bp}\selectfont\xeCJKsetup{PunctStyle=quanjiao,CJKglue=\hskip-0.06bp}\noindent\mbox{\xeCJKsetup{PunctStyle=plain}（}目录采用word 的自动生成目录方式，不能自己添加页码。）\par\vspace{2.4bp}}\fi\@starttoc{toc}\clearpage\endgroup}
\newcommand{\thesismainmatter}{\clearpage\@mainmattertrue\ecjtuwordgridtrue\pagenumbering{arabic}\pagestyle{thesis}}
\newcommand{\thesisbackchapter}[1]{\chapter*{#1}\addcontentsline{toc}{chapter}{#1}}
\newcommand{\thesisreferenceguidance}{%
  \ifecjtuguide
    {\color{red}\songti\fontsize{10.56bp}{19.5bp}\selectfont
    \xeCJKsetup{PunctStyle=quanjiao,CJKglue=\hskip-0.06bp}\noindent
    \mbox{\xeCJKsetup{PunctStyle=plain}（}参考文献书写注意以下几点：1\hspace{5.28bp}\textbf{参考文献正文}内容用\textbf{5号宋体字}，书写格式按照国家标准GB7714-87\linebreak
    规定，或者按照学生\textbf{毕业设计手册}\hspace{5.28bp}上关于参考文献的规定。\hspace{10.56bp}2\hspace{5.28bp}华东交通大学毕业设计要求参考文献\linebreak
    数量在\textbf{30篇以上}，近5年文献不少于70\%，其中至少有\textbf{3篇以上英文参考文献}。）\par\vspace{38.94bp}}
  \fi}
\renewcommand\bibsection{\thesisbackchapter{参考文献}\thesisreferenceguidance}
\renewcommand\bibfont{\ifecjtuguide\color{blue}\fi\songti\mdseries\fontsize{10.56bp}{19.5bp}\selectfont\xeCJKsetup{CJKglue=\hskip-0.06bp}}
\setlength\bibsep{0bp}
\bibliographystyle{gbt7714-numerical}
\newcommand{\ecjtuappendixlabel}[1]{{\XeTeXinterchartokenstate=0 附录\hspace{3.96bp}%
  {\rmfamily\mdseries\Alph{chapter}}}}
\newcommand{\thesisappendices}{\appendix\ctexset{chapter={name={附录,},number=\Alph{chapter},nameformat=\ecjtuappendixlabel,
  format=\centering\heiti\ecjtutitlelatin\mdseries\fontsize{15.96bp}{31.2bp}\selectfont\ecjtuwordstrut{21.4}{9.8}%
    \hspace*{\ifnum\value{chapter}=2 -0.28bp\else-0.36bp\fi}}}}
\AtBeginDocument{\hypersetup{pdftitle={\ecjtuTitle},pdfauthor={\ecjtuAuthor},pdfsubject={华东交通大学本科毕业论文}}}
% 把 TeX 实际采用的尺寸写出，供 scripts/check_format.py 与 PDF 交叉核验。
\newwrite\ecjtuMetrics
\AtBeginDocument{%
  \immediate\openout\ecjtuMetrics=\jobname.layout
  \immediate\write\ecjtuMetrics{paperwidth=\the\paperwidth}
  \immediate\write\ecjtuMetrics{paperheight=\the\paperheight}
  \immediate\write\ecjtuMetrics{left=\the\dimexpr1in+\hoffset+\oddsidemargin\relax}
  \immediate\write\ecjtuMetrics{evenleft=\the\dimexpr1in+\hoffset+\evensidemargin\relax}
  \immediate\write\ecjtuMetrics{right=\the\dimexpr\paperwidth-1in-\hoffset-\oddsidemargin-\textwidth\relax}
  \immediate\write\ecjtuMetrics{top=\the\dimexpr1in+\voffset+\topmargin+\headheight+\headsep\relax}
  \immediate\write\ecjtuMetrics{bottom=\the\dimexpr\paperheight-1in-\voffset-\topmargin-\headheight-\headsep-\textheight\relax}
  \immediate\write\ecjtuMetrics{header=\the\dimexpr1in+\voffset+\topmargin\relax}
  \immediate\write\ecjtuMetrics{footer=\the\dimexpr\paperheight-1in-\voffset-\topmargin-\headheight-\headsep-\textheight-\footskip\relax}
  \immediate\write\ecjtuMetrics{bodyfont=\f@size pt}
  \immediate\write\ecjtuMetrics{baseline=\the\baselineskip}
  \immediate\closeout\ecjtuMetrics}
\endinput
