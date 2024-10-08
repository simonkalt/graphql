/*
嬪様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様邑
� Source file: OBSOLETE.CH                                                 額
� Description: Conversions for obsolete SIx Driver functions.              額
� Notes      : This file is referenced by all of the other SIx Driver      額
�              header files.                                               額
� Last Update: 08/18/94                                                    額
� Notice     : Copyright 1994 - SuccessWare 90, Inc.                       額
塒様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様余�
  栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩栩�
*/

#xtranslate Sx_IndexFilter( [<nOrder>] )           => OrdFor( [<nOrder>] )
#xtranslate Sx_TagName([<nOrder>])                 => iif( empty(alias()), "",OrdName([<nOrder>]))
#xtranslate Sx_TagOrder()                          => iif( empty(alias()), 0 ,OrdNumber() )
#xtranslate Sx_SetTagOrder( [<xOrder>] [,<cBag>] ) => Sx_SetTag( [<xOrder>] [,<cBag>]  )
#xtranslate Sx_SetTagNo( <xOrder> )                => Sx_SetTag( <xOrder>, iif(ordNumber()>0,ordBagName(),ordBagName(1)))
#xtranslate Sx_SetTagNo( <xOrder>, <cBag> )        => Sx_SetTag( <xOrder>, <cBag> )
#xtranslate _sxCondSet( [<params,...>] )           => OrdCondSet( [<params>] )
#xtranslate SetRDD( [<cRDDname>] )                 => RDDSetDefault( [<cRDDname>] )

