/*컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴커
�                     Ganahl Software, Inc.                      �
�                        (305) 484-8394                          �
�                                                                �
�          (c) 1990-1994 Copyright Ganahl Software, Inc.         �
�                      All Rights Reserved                       �
�                                                                �
읕컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴*/
/*컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴커
�   The #define CLIP501 or CLIP52 should really be done on the   �
�   command line using the /D switch of the Clipper compiler.    �
�   However, because some people prefer to have things setup     �
�   in a more permanent fashion, you may uncomment one or the    �
�   other of the following #defines to make default behavior.    �
읕컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴*/
//#define CLIP501   // Uncomment this line to use Clipper 5.01a
#define CLIP52    // Uncomment this line to use Clipper 5.2
#ifndef CLIP501
   #ifndef CLIP52
      #error 'CLIP501' or 'CLIP52' must be #defined
      #error An exmple of making a #define from the
      #error command line switch is:
      #error    CLIPPER filename /b /m /w /DCLIP52
   #endif
#endif

#ifndef FLEXRDD_
   #define FLEXRDD_

   #ifdef CLIP501
      EXTERNAL _VFLEXFILE
      #command COPY [TO <(file)>]                                           ;
               [FIELDS <fields,...>]                                        ;
               [FOR <for>]                                                  ;
               [WHILE <while>]                                              ;
               [NEXT <next>]                                                ;
               [RECORD <rec>]                                               ;
               [<rest:REST>]                                                ;
               [VIA <rdd>]                                                  ;
               [ALL]                                                        ;
                                                                            ;
            => __flexCopy(                                                  ;
                         <(file)>, { <(fields)> },                          ;
                         <{for}>, <{while}>, <next>, <rec>, <.rest.>, <rdd> ;
                         )

      #command APPEND [FROM <(file)>]                                       ;
               [FIELDS <fields,...>]                                        ;
               [FOR <for>]                                                  ;
               [WHILE <while>]                                              ;
               [NEXT <next>]                                                ;
               [RECORD <rec>]                                               ;
               [<rest:REST>]                                                ;
               [VIA <rdd>]                                                  ;
               [ALL]                                                        ;
                                                                            ;
            => __flexApp(                                                   ;
                          <(file)>, { <(fields)> },                         ;
                          <{for}>, <{while}>, <next>, <rec>, <.rest.>, <rdd>;
                        )

      #command CREATE <(file1)>                                             ;
                  FROM <(file2)>                                            ;
                  [VIA <rdd>]                                               ;
                  [ALIAS <a>]                                               ;
                  [<new: NEW>]                                              ;
                                                                            ;
            => __flexCreate( <(file1)>, <(file2)>, <rdd>, <.new.>, <(a)> )


      #xtranslate dbcreate( <args,...> ) => v_dbfCreate( <args> )

   #else
      REQUEST FLEXFILE
   #endif

#endif
