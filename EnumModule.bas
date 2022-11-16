Attribute VB_Name = "EnumModule"
'The MIT License (MIT)
'
'Copyright (c) 2022 FORREST
' Mateusz Milewski mateusz@stellantis.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.



' static coupling - always double check
' ==============================================================

' zero will default
Public Enum EMOD
    emodRectCorail = 1
    emodRectWhite
    emodRectCorail1355
    emodRectCorail1320
    emodRectCorailNN
    emoddecorator
    emodpcv
    emodlook
End Enum


Public Enum eRecDataStd
    eRDS__NOK = 0
    eRDS__Corail1355 = EMOD.emodRectCorail1355
    eRDS__Corail1320 = EMOD.emodRectCorail1320
    eRDS__CorailNN = EMOD.emodRectCorailNN
    eRDS__PCV_std_export1 = EMOD.emodpcv
    eRDS__LOOK_std_export1 = EMOD.emodlook
End Enum


' Eif stands for E_(WHICH INPUT)FILE - what standard of the opened file I expect
Public Enum eWhichInputFile
    EifCorail1355
    EifCorail1320
    EifCorailNN
    EifCorailCustom
    EifPCV001
    EifPCV002
    EifLOOK001
    EifLOOK002
    EifCustom
End Enum

' ==============================================================




Public Enum EEMO
    l_echeance_est_fixée = 1
    GAcNomNOA
    Article
    designation
    fourn
    nomfournisseur
    RU
    affaire
    dateecheance
    teecheancée
    quantitelivree
    Datelivr
    Mag
    datedepassagePegase
    NOA
    Sousprojet
    Docachat
End Enum






Public Enum eInDecoratorInScenario
    ESCORAIL = 1
    ESPCV
    ESLOOK
End Enum

Public Enum eOutDecoratorOutScenario
    ESPUS = 2
    ESPROCURE
    ESTEARDOWN
    ESPUSPROCURE
    ESPUSTEARDOWN
End Enum
