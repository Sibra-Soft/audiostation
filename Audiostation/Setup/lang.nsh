!macro LanguageCodeToText LANG OUTVAR
    ${Select} ${LANG}
        ${Case} 1033
            StrCpy ${OUTVAR} "English"
        ${Case} 1043
            StrCpy ${OUTVAR} "Dutch"
        ${Case} 1031
            StrCpy ${OUTVAR} "German"
        ${Default}
            StrCpy ${OUTVAR} "English"
    ${EndSelect}
!macroend