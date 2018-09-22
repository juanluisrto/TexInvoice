\NeedsTeXFormat{LaTeX2e} 
\ProvidesClass{faktura}
\LoadClass[12pt]{article}

\usepackage[letterpaper,hmargin=0.79in,vmargin=0.79in]{geometry}
\usepackage[parfill]{parskip} % Do not indent paragraphs
\usepackage{fp} % Fixed-point arithmetic
\usepackage[utf8]{inputenc}
%\usepackage[swedish]{babel}
\usepackage{longtable}

\pagestyle{empty} % No page numbers
\linespread{1.5} % Line spacing
\setlength{\doublerulesep}{\arrayrulewidth} % Double rules look like one thick one



\FPset{\result}{0}
\FPset{\moms}{0}

\def \tab {\hspace*{3ex}} % Define \tab to create some horizontal white space

\newcommand{\printStiftelseProCultura}{%
	{\LARGE \bf Stiftelsen Pro Cultura} \\ % Company providing the invoice
	
	\hfil{\Large\bf Faktura}\hfil % Company providing the invoice
	\bigskip\break % Whitespace
	\hrule % Horizontal line
	Friggagatan 10 \hfill +46 8 20 42 50 \\ % Your address and contact information
	11427, Stockholm \hfill stiftelsen@procultura.se \\ \\}

\newcommand*{\printReceiver}[1]{%
	{\bf Mottagare:} \\%
		\tab #1 \\ \\ } % Invoice recipient


\newcommand*{\printDates}[3]{
	{\bf Fakturanummer:} #1 \\ % Invoice number
	{\bf Fakturadatum:}  #2 \\ % Invoice date
	{\bf Förfallodatum:} #3 \\ % Expiry date
}
	
% Formats inputed number with 2 digits after the decimal place
\newcommand*{\formatNumber}[1]{\FPround{\cost}{#1}{2}\cost} %  


\newcommand{\LarkstadenInfo}{
	{\tab  Org n: 802015-9755} \\
	{\tab  Pg: 709896-5} \\
}

\newcommand{\ProCulturaInfo}{
	{\tab  Org n: 802015-9755} \\
	{\tab  Pg: XXXXXX-X} \\
}

\newenvironment{invoiceTable}{

	% Create a new row from title and expense amount
	\newcommand*{\feerow}[2]{%
        \FPadd{\result}{\result}{##2}
        \xdef\result{\result}
        ##1 & \formatNumber{##2} kr \\
        }

    % Create a row with the percentage of vat (moms) which is not included in the final sum
    \newcommand*{\momsrow}[1]{%
        \FPmul{\moms}{\result}{##1}
        \FPdiv{\moms}{\moms}{100}
        \xdef\moms{\moms}
        \footnotesize{Moms ##1\% ingår med} & \footnotesize{\formatNumber{\moms} kr} \\
        }

    \setlength\LTleft{0pt}
    \setlength\LTright{0pt}
    \begin{longtable}{@{\extracolsep{\fill}\hspace{\tabcolsep}} l r }
    \hline
    {\bf Beskrivning} & \multicolumn{1}{c}{\bf Summa} \\*
    \hline\hline
    \endhead
}
{   
    \hline\hline\hline
    {\bf Att betala:}  & {\bf \formatNumber{\result} kr} \\
    \end{longtable}
}

