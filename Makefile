HEADERS = microxlswriter.h
OBJECTS = microxlswriter.o test.o

default: test

%.o: %.c $(HEADERS)
	gcc -Os -c $< -o $@

test: $(OBJECTS)
	gcc $(OBJECTS) -o $@

clean:
	-rm -f $(OBJECTS)
	-rm -f test
