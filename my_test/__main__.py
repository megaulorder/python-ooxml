from my_test import app
import time

if __name__ == '__main__':
    start_time = time.time()
    app.run()
    print('\n--- %s seconds ---' % (time.time() - start_time))
