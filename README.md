# Общие слова

Скрипт готов к работе. Я написал его, чтобы запускать несколько раз вручную, самостоятельно. Выходные данные сейчас выводятся в консоль. При необходимости, могу дописать вывод данных в .xls файлы.

Для работы скрипта нужен файлик "rcb_uurraa_data.xlsx". Он есть в репозитории. Коэффициенты задаются в вкладке "coefficients". Данные скрипт забирает из вкладки "main".

По всем вопросам пишите на tkorchagin@gmail.com

# Алгоритм AHP

## Основная идея
Для ранжирования регионов использую метод аналитической иерархии. [Здесь](https://yadi.sk/i/W7gnlVaBmzDEG) отлично расписан алгоритм. Советую.

Я расписал своими словами.

По каждому критерию находим сколько весит каждый регион. Потом находим вес каждого критерия. Потом для каждого региона вычисляем его полезность: сумма произведений веса критерия на вес региона по этому критерию.

## По пунктам
### 1. Считываем данные
В файлике есть столбцы. Каждая строка -- это величина определенного критерия для конкретного региона. Столбец -- это величины конкретного критерия по всем регионам.

Достаем данные из файлика. Выбираем столбик. Считываем столбик чисел и его название.

    def get_data(fn, sheet_name, column, from_row=0):
        wb = load_workbook(fn)

        sheet = wb.get_sheet_by_name(sheet_name)
        row_count = sheet.get_highest_row()

        column_name = sheet.cell(row=0, column=column).value

        data = []
        for q in xrange(1, row_count):
            if q < from_row:
                continue
            cell = sheet.cell(row=q, column=column)
            value = cell.value
            data.append(value)
        return data, column_name

По названию столбика решаем, что лучше: чем величина критерия больше или меньше.

### 2. Нормализуем данные

    def normalise(arr):
        # for item in arr:
        #     print item
        square_arr = [item ** 2 for item in arr]
        square_sum = sum(square_arr)
        norm_coefficient = square_sum ** 0.5

        if norm_coefficient == 0:
            return None

        norm_arr = [item / norm_coefficient for item in arr]
        return norm_arr

### 3. Вес региона по критерию
Когда решили, что лучше, большой критерий или маленький, получаем **вес региона по конкретному критерию.**

    def get_weights(data, bigger_better):
        norma = normalise(data)
        self_vectors = []
        for q in xrange(len(norma)):
            preference_arr = degree_of_preference(q, norma, bigger_better)
            s_vector = self_vector(preference_arr)
            self_vectors.append(s_vector)
        weights = alternative_weight(self_vectors)
        return weights

### 4. Относительная важность
Для вычисления веса региона во первых нужно вычислить **относительную важность** региона по этому критерию.

    def degree_of_preference(index, arr, bigger_better=True):
        result = []
        value = arr[index] + 0.05  # Защита от деления на 0
        for q in xrange(len(arr)):
            div = arr[q] + 0.05  # Защита от деления на 0
            if bigger_better:
                v = value / div
            else:
                v = div / value
            result.append(v)
        return result

### 5. Собственный вектор
Теперь вычисляем **собственный вектор** по этому критерию для каждого региона.

    def self_vector(arr):
        mult = 1
        for el in arr:
            mult *= el

        power = 1.0 / len(arr)
        return mult ** power

### 6. Вес региона по критерию
В результате нормализации собственных векторов вычисляем веса регионов по каждому критерию и веса самих критериев.

    def alternative_weight(arr):
        result = [float(item) / sum(arr) for item in arr]
        return result

Также вычисляются и **веса критериев.**
### 7. Полезность региона
В итоге считаем полезность конкретного региона. Для этого суммируем произведения веса критерия на вес конкретного региона по этому кртерию.

### P.S.
Для того, чтобы алгоритм работал, надо иметь полностью заполненные столбики. Пропусков быть не должно.
