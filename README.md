Autotyper para facilitar la creación de roadcases en RentalPoint.

Hasta el 10/04/2019, RP no permite crear varias roadcases en una sola acción, ni tampoco comprueba que los productos que se ponen dentro de un roadcase sean los correctos. 
Por eso se creó este programa.

Este programa no consulta a la base de datos de RP, por lo que no comprueba si los códigos están en la location correcta, ni si el roadcase ya está creado ni si los assets no se han retornado.
Hasta el 10/04/2019, RP solo avisa cuando se añade en un Roadcase un producto que ya está en otro Roadcase.

La creación de nuevos paquetes hay que hacerla mediante código, pues no hay creado un front-end ni base de datos para hacerlo.
Normalmente es suficiente con copiar el bloque del paquete que tenga más componentes, modificar el nombre de las variables para facilitar su mantenimiento, modificar los mensajes y borrar las variables que sobren.

Una vez compilado el programa, copiar el .exe al escritorio remoto donde se use RP.
Hasta el 10/04/2019 hay una carpeta Helpers en la unidad de red de Almacén para que todos los remotos de almacén tenga acceso a este y otros archivos de ayuda.
