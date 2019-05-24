Autotyper para facilitar la creación de roadcases en RentalPoint.

Hasta el 10/04/2019, RP no permite crear varias roadcases en una sola acción, ni tampoco comprueba que los productos que se ponen dentro de un roadcase sean los correctos.
Por eso se creó este programa.

Este programa no consulta a la base de datos de RP, por lo que no comprueba si los códigos están en la location correcta, ni si el roadcase ya está creado ni si los assets no se han retornado.
Hasta el 10/04/2019, RP solo avisa cuando se añade en un Roadcase un producto que ya está en otro Roadcase.

La declaración de paquetes se hace en el archivo txt "Roadcase Creator Database.txt".
Incluyo una copia en la raiz de este repositorio.
Este archivo tiene que estar en la misma ruta que el ejecutable.

Una vez compilado el programa, copiar el .exe (y el txt anteriormente citado si no lo tienes ya, para que te sirba de guia) al escritorio remoto donde se use RP.
Hasta el 10/04/2019 hay una carpeta Helpers en la unidad de red de Almacén para que todos los remotos de almacén tenga acceso a este y otros archivos de ayuda.
